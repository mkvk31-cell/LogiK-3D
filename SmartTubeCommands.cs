using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace LogiK3D.Piping
{
    public class SmartTubeCommands
    {
        // Constantes de configuration
        private const string AppName = "LogiK_Data";

        [CommandMethod("LOGIK_PIPE")]
        public void CreateSmartTube()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Demander le numéro de ligne
            PromptResult prLine = ed.GetString("\nEntrez le numéro de la ligne (ex: L100) : ");
            if (prLine.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(prLine.StringResult))
            {
                ed.WriteMessage("\nNuméro de ligne invalide. Commande annulée.");
                return;
            }
            string lineNumber = prLine.StringResult;

            // Récupérer les paramètres depuis la palette
            double currentOD = LogiK3D.UI.MainPaletteControl.CurrentOuterDiameter;
            string currentDN = LogiK3D.UI.MainPaletteControl.CurrentDN;
            double currentThickness = LogiK3D.UI.MainPaletteControl.CurrentThickness;

            // 1. Sélection des polylignes
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nSélectionnez les axes des tubes (Polylignes ou Polylignes 3D) : ";
            
            // Filtre pour ne sélectionner que les polylignes
            TypedValue[] filterList = new TypedValue[] {
                new TypedValue((int)DxfCode.Operator, "<OR"),
                new TypedValue((int)DxfCode.Start, "POLYLINE"),
                new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                new TypedValue((int)DxfCode.Start, "POLYLINE3D"),
                new TypedValue((int)DxfCode.Operator, "OR>")
            };
            SelectionFilter filter = new SelectionFilter(filterList);

            PromptSelectionResult psr = ed.GetSelection(pso, filter);
            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nCommande annulée ou sélection vide.");
                return;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    foreach (SelectedObject so in psr.Value)
                    {
                        Entity ent = (Entity)tr.GetObject(so.ObjectId, OpenMode.ForWrite);
                        
                        // Attacher les données de ligne à la polyligne
                        PipeManager.AttachLineData(ent, lineNumber, currentDN, currentOD, currentThickness, tr, db);

                        // Générer la tuyauterie
                        PipeManager.GeneratePiping(ent, lineNumber, currentDN, currentOD, currentThickness, tr);
                    }

                    tr.Commit();
                    ed.WriteMessage($"\nRéseau LogiK 3D (Ligne {lineNumber}) généré avec succès sur {psr.Value.Count} polyligne(s).");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nErreur lors de la génération : {ex.Message}");
                    tr.Abort();
                }
            }
        }

        [CommandMethod("LOGIK_INSERT_COMP")]
        public void InsertComponent()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Récupérer le type de composant depuis les arguments de la commande
            // (AutoCAD passe les arguments via LISP, mais on peut aussi demander à l'utilisateur)
            PromptResult prType = ed.GetString("\nType de composant (BRIDE, COUDE, TEE, REDUCER) : ");
            if (prType.Status != PromptStatus.OK) return;
            
            string compType = prType.StringResult.ToUpper();

            // Récupérer les paramètres depuis la palette
            double currentOD = LogiK3D.UI.MainPaletteControl.CurrentOuterDiameter;
            string currentDN = LogiK3D.UI.MainPaletteControl.CurrentDN;
            double currentThickness = LogiK3D.UI.MainPaletteControl.CurrentThickness;

            string targetDN = currentDN;
            double targetOD = currentOD;

            if (compType == "REDUCER" || compType == "RED_CONC")
            {
                PromptStringOptions psoDN = new PromptStringOptions($"\nEntrez le DN de réduction (ex: DN80) [Actuel: {currentDN}] : ");
                psoDN.AllowSpaces = false;
                PromptResult prDN = ed.GetString(psoDN);
                if (prDN.Status != PromptStatus.OK) return;

                string inputDN = prDN.StringResult.ToUpper();
                if (!inputDN.StartsWith("DN")) inputDN = "DN" + inputDN;

                if (LogiK3D.UI.MainPaletteControl.AvailableDiameters != null && 
                    LogiK3D.UI.MainPaletteControl.AvailableDiameters.ContainsKey(inputDN))
                {
                    targetDN = inputDN;
                    targetOD = LogiK3D.UI.MainPaletteControl.AvailableDiameters[inputDN];
                }
                else
                {
                    ed.WriteMessage($"\nErreur : Le diamètre {inputDN} n'est pas reconnu dans la spécification.");
                    return;
                }
            }

            // Demander le point d'insertion
            PromptPointOptions ppo = new PromptPointOptions($"\nSpécifiez le point d'insertion pour {compType} {currentDN}{(compType.StartsWith("RED") ? " vers " + targetDN : "")} : ");
            PromptPointResult ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK) return;

            Point3d insertPt = ppr.Value;

            // Demander la direction (rotation)
            PromptPointOptions ppoDir = new PromptPointOptions("\nSpécifiez la direction : ");
            ppoDir.UseBasePoint = true;
            ppoDir.BasePoint = insertPt;
            PromptPointResult pprDir = ed.GetPoint(ppoDir);
            if (pprDir.Status != PromptStatus.OK) return;

            Vector3d direction = (pprDir.Value - insertPt).GetNormal();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Si c'est une réduction, on va essayer de couper la polyligne existante
                    if (compType == "REDUCER" || compType == "RED_CONC")
                    {
                        // 1. Trouver la polyligne sous le point d'insertion
                        PromptSelectionOptions pso = new PromptSelectionOptions();
                        pso.MessageForAdding = "\nSélectionnez la ligne (Polyligne) à couper pour insérer la réduction : ";
                        
                        TypedValue[] filterList = new TypedValue[] {
                            new TypedValue((int)DxfCode.Operator, "<OR"),
                            new TypedValue((int)DxfCode.Start, "POLYLINE"),
                            new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                            new TypedValue((int)DxfCode.Start, "POLYLINE3D"),
                            new TypedValue((int)DxfCode.Operator, "OR>")
                        };
                        SelectionFilter filter = new SelectionFilter(filterList);

                        PromptSelectionResult psr = ed.GetSelection(pso, filter);
                        if (psr.Status == PromptStatus.OK && psr.Value.Count > 0)
                        {
                            ObjectId polyId = psr.Value[0].ObjectId;
                            Entity polyEnt = (Entity)tr.GetObject(polyId, OpenMode.ForWrite);
                            
                            // Récupérer les données de la ligne d'origine
                            string lineNumber = "";
                            ResultBuffer rb = polyEnt.GetXDataForApplication(PipeManager.LineDataAppName);
                            if (rb != null)
                            {
                                TypedValue[] values = rb.AsArray();
                                if (values.Length >= 2) lineNumber = values[1].Value.ToString();
                            }

                            if (polyEnt is Polyline poly)
                            {
                                // Trouver le segment le plus proche du point d'insertion
                                Point3d closestPt = poly.GetClosestPointTo(insertPt, false);
                                double param = poly.GetParameterAtPoint(closestPt);
                                int segmentIndex = (int)Math.Floor(param);
                                
                                // Créer deux nouvelles polylignes
                                Polyline poly1 = new Polyline();
                                Polyline poly2 = new Polyline();
                                
                                // Remplir poly1 (du début jusqu'au point de coupure)
                                int v1 = 0;
                                for (int i = 0; i <= segmentIndex; i++)
                                {
                                    poly1.AddVertexAt(v1++, poly.GetPoint2dAt(i), poly.GetBulgeAt(i), poly.GetStartWidthAt(i), poly.GetEndWidthAt(i));
                                }
                                poly1.AddVertexAt(v1, new Point2d(closestPt.X, closestPt.Y), 0, 0, 0);
                                
                                // Remplir poly2 (du point de coupure jusqu'à la fin)
                                int v2 = 0;
                                poly2.AddVertexAt(v2++, new Point2d(closestPt.X, closestPt.Y), 0, 0, 0);
                                for (int i = segmentIndex + 1; i < poly.NumberOfVertices; i++)
                                {
                                    poly2.AddVertexAt(v2++, poly.GetPoint2dAt(i), poly.GetBulgeAt(i), poly.GetStartWidthAt(i), poly.GetEndWidthAt(i));
                                }

                                BlockTableRecord btrSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                                
                                // Ajouter poly1 (Garde le diamètre d'origine)
                                btrSpace.AppendEntity(poly1);
                                tr.AddNewlyCreatedDBObject(poly1, true);
                                PipeManager.AttachLineData(poly1, lineNumber, currentDN, currentOD, currentThickness, tr, db);
                                
                                // Ajouter poly2 (Prend le nouveau diamètre réduit)
                                btrSpace.AppendEntity(poly2);
                                tr.AddNewlyCreatedDBObject(poly2, true);
                                // On suppose que l'épaisseur reste la même pour simplifier, ou on pourrait la chercher
                                PipeManager.AttachLineData(poly2, lineNumber, targetDN, targetOD, currentThickness, tr, db);

                                // Supprimer l'ancienne polyligne et ses solides
                                PipeManager.DeleteLinkedSolids(poly, tr);
                                poly.Erase();
                                
                                ed.WriteMessage($"\nLigne coupée. La suite de la ligne passe en {targetDN}.");
                                
                                // Mettre à jour le point d'insertion et la direction pour s'aligner parfaitement sur la ligne
                                insertPt = closestPt;
                                if (segmentIndex + 1 < poly.NumberOfVertices)
                                {
                                    Point3d nextPt = poly.GetPoint3dAt(segmentIndex + 1);
                                    direction = (nextPt - closestPt).GetNormal();
                                }
                            }
                            else if (polyEnt is Polyline3d poly3d)
                            {
                                ed.WriteMessage("\nLa coupure automatique sur Polyline3D n'est pas encore supportée. Le bloc sera inséré normalement.");
                            }
                        }
                    }

                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    PipingGenerator generator = new PipingGenerator();
                    ObjectId blockId = ObjectId.Null;

                    if (compType == "BRIDE")
                    {
                        // Chercher les données dans la base Kohler
                        var flangeData = LogiK3D.Specs.KohlerDatabase.FindFlange(currentDN);
                        
                        double flangeOD = flangeData != null && flangeData.OuterDiameter > 0 ? flangeData.OuterDiameter : currentOD * 1.5; // Valeur par défaut si non trouvé
                        double flangeThickness = flangeData != null && flangeData.WallThickness > 0 ? flangeData.WallThickness : 16.0; // Valeur par défaut
                        double flangeLength = flangeThickness + 30.0; // Longueur totale avec collet
                        
                        string blockName = $"FLANGE_{currentDN}_{flangeOD}_{flangeThickness}";
                        blockId = generator.GetOrCreateFlange(currentOD, flangeOD, flangeThickness, flangeLength, blockName);
                        
                        if (flangeData != null)
                        {
                            ed.WriteMessage($"\nDonnées Kohler trouvées : {flangeData.Type} (OD: {flangeData.OuterDiameter}, Ep: {flangeData.WallThickness})");
                        }
                    }
                    else if (compType == "COUDE")
                    {
                        double radius = currentOD * 1.5; // Rayon 3D standard
                        string blockName = $"ELBOW_90_{currentDN}_{radius}";
                        blockId = generator.GetOrCreateElbow(currentOD, radius, 90.0, blockName);
                    }
                    else if (compType == "TEE")
                    {
                        double length = currentOD * 2.0;
                        double branchHeight = currentOD;
                        string blockName = $"TEE_{currentDN}";
                        blockId = generator.GetOrCreateTee(currentOD, currentOD, length, branchHeight, blockName);
                    }
                    else if (compType == "REDUCER" || compType == "RED_CONC")
                    {
                        double length = currentOD * 1.5; // Longueur standard
                        string blockName = $"REDUCER_{currentDN}_{targetDN}";
                        blockId = generator.GetOrCreateReducer(currentOD, targetOD, length, blockName);
                    }
                    else
                    {
                        ed.WriteMessage($"\n[DEBUG] Type de composant non reconnu par le générateur : '{compType}'");
                    }

                    if (blockId != ObjectId.Null)
                    {
                        // Insérer la référence de bloc
                        BlockReference bref = new BlockReference(insertPt, blockId);
                        
                        // Aligner le bloc avec la direction
                        Vector3d xAxis = Vector3d.XAxis;
                        double angle = xAxis.GetAngleTo(direction);
                        Vector3d axisOfRotation = xAxis.CrossProduct(direction);
                        
                        if (!axisOfRotation.IsZeroLength())
                        {
                            bref.TransformBy(Matrix3d.Rotation(angle, axisOfRotation, insertPt));
                        }
                        else if (direction.X < 0)
                        {
                            bref.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, insertPt));
                        }

                        btr.AppendEntity(bref);
                        tr.AddNewlyCreatedDBObject(bref, true);
                        
                        ed.WriteMessage($"\n{compType} inséré avec succès.");

                        // Si c'est une réduction, on met à jour la palette pour que le prochain tube soit à la bonne taille
                        if (compType == "REDUCER" || compType == "RED_CONC")
                        {
                            if (LogiK3D.UI.MainPaletteControl.Instance != null)
                            {
                                LogiK3D.UI.MainPaletteControl.Instance.SetCurrentDN(targetDN);
                                ed.WriteMessage($"\n[Info] Le diamètre actif de la palette a été mis à jour sur {targetDN}.");
                            }
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"\nErreur : Impossible de générer le bloc pour {compType}.");
                    }

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nErreur fatale lors de la génération du bloc : {ex.Message}\n{ex.StackTrace}");
                    tr.Abort();
                }
            }
        }

        [CommandMethod("LOGIK_BOM")]
        public void GenerateIsoBOM()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n--- NOMENCLATURE LOGIK 3D (ISO) ---");
            ed.WriteMessage("\n{0,-15} | {1,-20} | {2,-15} | {3,-20}", "Ligne", "Composant", "Longueur (mm)", "Code SAP Kohler");
            ed.WriteMessage("\n-------------------------------------------------------------------------------");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                int count = 0;

                foreach (ObjectId id in btr)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent != null && ent.XData != null)
                    {
                        ResultBuffer rb = ent.GetXDataForApplication(PipeManager.SolidDataAppName);
                        if (rb != null)
                        {
                            TypedValue[] values = rb.AsArray();
                            // Structure attendue : [0] RegApp, [1] GUID, [2] SAP Code, [3] Length, [4] Line Number
                            if (values.Length >= 5)
                            {
                                string sapCode = values[2].Value.ToString();
                                double length = (double)values[3].Value;
                                string lineNumber = values[4].Value.ToString();

                                ed.WriteMessage($"\n{lineNumber,-15} | {ent.GetType().Name,-20} | {length,-15:F2} | {sapCode,-20}");
                                count++;
                            }
                        }
                    }
                }

                ed.WriteMessage("\n-------------------------------------------------------------------------------");
                ed.WriteMessage($"\nTotal composants trouvés : {count}\n");
                tr.Commit();
            }
        }
    }
}