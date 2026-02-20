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

            // Récupérer les paramètres depuis la palette
            double currentOD = LogiK3D.UI.MainPaletteControl.CurrentOuterDiameter;
            double pipeRadius = currentOD / 2.0;
            
            // Extraire la valeur numérique du DN (ex: "DN100" -> 100)
            string dnString = LogiK3D.UI.MainPaletteControl.CurrentDN.Replace("DN", "");
            double dnValue = 100.0;
            double.TryParse(dnString, out dnValue);
            
            // Rayon standard 1.5D pour les coudes
            double elbowRadiusLR = 1.5 * dnValue;

            // 1. Sélection de la polyligne
            PromptEntityOptions peo = new PromptEntityOptions("\nSélectionnez l'axe du tube (Polyligne ou Polyligne 3D) : ");
            peo.SetRejectMessage("\nL'objet doit être une polyligne.");
            peo.AddAllowedClass(typeof(Polyline), exactMatch: false);
            peo.AddAllowedClass(typeof(Polyline3d), exactMatch: false);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nCommande annulée ou sélection vide.");
                return;
            }

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    Entity ent = (Entity)tr.GetObject(per.ObjectId, OpenMode.ForRead);
                    List<Point3d> vertices = GetPolylineVertices(ent, tr);

                    if (vertices.Count < 2)
                    {
                        ed.WriteMessage("\nLa polyligne doit contenir au moins 2 sommets.");
                        return;
                    }

                    EnsureRegAppExists(db, tr, AppName);

                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    double[] cutbacks = new double[vertices.Count];
                    
                    // 2. Calcul des angles et insertion des coudes
                    for (int i = 1; i < vertices.Count - 1; i++)
                    {
                        Vector3d vIn = (vertices[i] - vertices[i - 1]).GetNormal();
                        Vector3d vOut = (vertices[i + 1] - vertices[i]).GetNormal();
                        
                        double angleRad = vIn.GetAngleTo(vOut);
                        double angleDeg = angleRad * (180.0 / Math.PI);

                        if (angleDeg > 1.0) // Ignorer les sommets quasi-alignés
                        {
                            // Calcul du retrait (Cutback) pour tronquer le tube
                            cutbacks[i] = Math.Tan(angleRad / 2.0) * elbowRadiusLR;

                            string blockName = angleDeg > 60.0 ? "KOHLER_ELBOW_90" : "KOHLER_ELBOW_45";

                            if (bt.Has(blockName))
                            {
                                InsertElbowBlock(vertices[i], vIn, vOut, blockName, bt, btr, tr);
                            }
                            else
                            {
                                ed.WriteMessage($"\nAttention: Bloc '{blockName}' introuvable dans le dessin.");
                            }
                        }
                    }

                    // 3. Génération des tubes (Solid3d) tronqués
                    for (int i = 0; i < vertices.Count - 1; i++)
                    {
                        Vector3d direction = (vertices[i + 1] - vertices[i]).GetNormal();
                        Point3d startPt = vertices[i] + direction * cutbacks[i];
                        Point3d endPt = vertices[i + 1] - direction * cutbacks[i + 1];

                        double cutLength = startPt.DistanceTo(endPt);

                        if (cutLength > 0)
                        {
                            Solid3d pipeSolid = CreatePipeSolid(startPt, endPt, pipeRadius);
                            
                            // Ajout des XData
                            AttachLogiKData(pipeSolid, $"KOH-{LogiK3D.UI.MainPaletteControl.CurrentDN}-PIPE", cutLength);

                            btr.AppendEntity(pipeSolid);
                            tr.AddNewlyCreatedDBObject(pipeSolid, true);
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage("\nRéseau LogiK 3D généré avec succès.");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nErreur lors de la génération : {ex.Message}");
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
            ed.WriteMessage("\n{0,-20} | {1,-15} | {2,-20}", "Composant", "Longueur (mm)", "Code SAP Kohler");
            ed.WriteMessage("\n---------------------------------------------------------------");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                int count = 0;

                foreach (ObjectId id in btr)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent != null && ent.XData != null)
                    {
                        ResultBuffer rb = ent.GetXDataForApplication(AppName);
                        if (rb != null)
                        {
                            TypedValue[] values = rb.AsArray();
                            // Structure attendue : [0] RegApp, [1] GUID, [2] SAP Code, [3] Length
                            if (values.Length >= 4)
                            {
                                string sapCode = values[2].Value.ToString();
                                double length = (double)values[3].Value;

                                ed.WriteMessage($"\n{ent.GetType().Name,-20} | {length,-15:F2} | {sapCode,-20}");
                                count++;
                            }
                        }
                    }
                }

                ed.WriteMessage("\n---------------------------------------------------------------");
                ed.WriteMessage($"\nTotal composants trouvés : {count}\n");
                tr.Commit();
            }
        }

        #region Méthodes Utilitaires

        private List<Point3d> GetPolylineVertices(Entity ent, Transaction tr)
        {
            List<Point3d> pts = new List<Point3d>();
            if (ent is Polyline poly)
            {
                for (int i = 0; i < poly.NumberOfVertices; i++)
                    pts.Add(poly.GetPoint3dAt(i));
            }
            else if (ent is Polyline3d poly3d)
            {
                foreach (ObjectId vId in poly3d)
                {
                    PolylineVertex3d v3d = (PolylineVertex3d)tr.GetObject(vId, OpenMode.ForRead);
                    pts.Add(v3d.Position);
                }
            }
            return pts;
        }

        private Solid3d CreatePipeSolid(Point3d startPt, Point3d endPt, double radius)
        {
            double length = startPt.DistanceTo(endPt);
            Vector3d direction = (endPt - startPt).GetNormal();
            Point3d midPt = startPt + (direction * (length / 2.0));

            Solid3d pipe = new Solid3d();
            pipe.CreateFrustum(length, radius, radius, radius);

            // Aligner le cylindre (créé par défaut sur l'axe Z) avec la direction du segment
            Vector3d zAxis = direction;
            Vector3d xAxis = zAxis.IsParallelTo(Vector3d.XAxis) ? Vector3d.YAxis : zAxis.CrossProduct(Vector3d.XAxis).GetNormal();
            Vector3d yAxis = zAxis.CrossProduct(xAxis).GetNormal();

            Matrix3d transform = Matrix3d.AlignCoordinateSystem(
                Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                midPt, xAxis, yAxis, zAxis);

            pipe.TransformBy(transform);
            return pipe;
        }

        private void InsertElbowBlock(Point3d position, Vector3d vIn, Vector3d vOut, string blockName, BlockTable bt, BlockTableRecord btr, Transaction tr)
        {
            BlockReference br = new BlockReference(position, bt[blockName]);
            
            // Calcul du plan du coude pour l'orientation
            Vector3d normal = vIn.CrossProduct(vOut).GetNormal();
            if (normal.Length > 0.01)
            {
                // Aligne l'axe X du bloc sur le tube entrant, et Z sur la normale du plan
                Matrix3d blockTransform = Matrix3d.AlignCoordinateSystem(
                    Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis,
                    position, vIn, normal.CrossProduct(vIn).GetNormal(), normal);
                
                br.TransformBy(blockTransform);
            }

            btr.AppendEntity(br);
            tr.AddNewlyCreatedDBObject(br, true);
        }

        private void EnsureRegAppExists(Database db, Transaction tr, string appName)
        {
            RegAppTable rat = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (!rat.Has(appName))
            {
                rat.UpgradeOpen();
                RegAppTableRecord ratr = new RegAppTableRecord { Name = appName };
                rat.Add(ratr);
                tr.AddNewlyCreatedDBObject(ratr, true);
            }
        }

        private void AttachLogiKData(Entity ent, string sapCode, double cutLength)
        {
            ResultBuffer rb = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, AppName),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, Guid.NewGuid().ToString()),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, sapCode),
                new TypedValue((int)DxfCode.ExtendedDataReal, cutLength)
            );
            ent.XData = rb;
        }

        #endregion
    }
}