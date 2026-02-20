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

        [CommandMethod("LOGIK_UPDATE_LINE")]
        public void UpdateLine()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityOptions peo = new PromptEntityOptions("\nS�lectionnez la ligne � mettre � jour : ");
            peo.SetRejectMessage("\nVeuillez s�lectionner une polyligne.");
            peo.AddAllowedClass(typeof(Polyline), true);
            peo.AddAllowedClass(typeof(Polyline3d), true);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Entity polyEnt = (Entity)tr.GetObject(per.ObjectId, OpenMode.ForWrite);
                string lineNumber = "LIGNE_INCONNUE";
                string baseDN = LogiK3D.UI.MainPaletteControl.CurrentDN;
                double baseOD = LogiK3D.UI.MainPaletteControl.CurrentOuterDiameter;
                double thickness = LogiK3D.UI.MainPaletteControl.CurrentThickness;

                ResultBuffer rb = polyEnt.GetXDataForApplication(PipeManager.LineDataAppName);
                if (rb != null)
                {
                    TypedValue[] values = rb.AsArray();
                    if (values.Length >= 2) lineNumber = values[1].Value.ToString();
                    if (values.Length >= 4) baseDN = values[3].Value.ToString();
                    if (values.Length >= 5) double.TryParse(values[4].Value.ToString(), out baseOD);
                    if (values.Length >= 6) double.TryParse(values[5].Value.ToString(), out thickness);
                }

                PipeManager.UpdateLineComponents(polyEnt, lineNumber, baseDN, tr);
                PipeManager.DeleteLinkedSolids(polyEnt, tr);
                PipeManager.GeneratePiping(polyEnt, lineNumber, baseDN, baseOD, thickness, tr);

                tr.Commit();
                ed.WriteMessage("\nLigne $lineNumber mise � jour avec succ�s.");
            }
        }

        [CommandMethod("LOGIK_INSERT_COMP")]
        public void InsertComponent()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // On utilise GetString pour pouvoir recevoir l'argument depuis la palette
            PromptStringOptions pso = new PromptStringOptions("\nType de composant (BRIDE, COUDE, TEE, REDUCER, VANNE, VALVE_GLOBE, VALVE_BALL, CHECK_VALVE, FILTRE, DEBIMETRE, DIAPHRAGME) : ");
            pso.AllowSpaces = false;
            PromptResult prType = ed.GetString(pso);
            if (prType.Status != PromptStatus.OK) return;
            
            string compType = prType.StringResult.ToUpper();

            // Valeurs par défaut depuis la palette
            double currentOD = LogiK3D.UI.MainPaletteControl.CurrentOuterDiameter;
            string currentDN = LogiK3D.UI.MainPaletteControl.CurrentDN;
            double currentThickness = LogiK3D.UI.MainPaletteControl.CurrentThickness;
            string lineNumber = "LIGNE_INCONNUE";

            Point3d insertPt = Point3d.Origin;
            Vector3d direction = Vector3d.XAxis;
            bool isOnPipe = false;
            ObjectId selectedPolyId = ObjectId.Null;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 1. Demander de sélectionner une conduite pour l'alignement magnétique
                PromptEntityOptions peo = new PromptEntityOptions("\nSélectionnez la conduite pour insérer le composant (ou Entrée pour placement libre) : ");
                peo.SetRejectMessage("\nVeuillez sélectionner une polyligne.");
                peo.AddAllowedClass(typeof(Polyline), true);
                peo.AddAllowedClass(typeof(Polyline3d), true);
                peo.AllowNone = true;

                PromptEntityResult per = ed.GetEntity(peo);

                    Entity polyEnt = null;
                    if (per.Status == PromptStatus.OK)
                    {
                        isOnPipe = true;
                        selectedPolyId = per.ObjectId;
                        polyEnt = (Entity)tr.GetObject(selectedPolyId, OpenMode.ForWrite);
                        
                        // Récupérer le numéro de ligne de la polyligne
                        ResultBuffer rb = polyEnt.GetXDataForApplication(PipeManager.LineDataAppName);
                        if (rb != null)
                        {
                            TypedValue[] values = rb.AsArray();
                            if (values.Length >= 2) lineNumber = values[1].Value.ToString();
                        }

                        PromptPointOptions ppo = new PromptPointOptions($"\nSpécifiez le point d'insertion sur la ligne pour {compType} {currentDN} : ");
                        PromptPointResult ppr = ed.GetPoint(ppo);
                        if (ppr.Status != PromptStatus.OK) return;

                        // Calculer le point d'insertion exact et la direction (tangente)
                        Curve curve = polyEnt as Curve;
                        if (curve != null)
                        {
                            insertPt = curve.GetClosestPointTo(ppr.Value, false);
                            double param = curve.GetParameterAtPoint(insertPt);
                            direction = curve.GetFirstDerivative(param).GetNormal();
                        }
                    }
                    else if (per.Status == PromptStatus.None)
                    {
                        // Placement libre
                        PromptPointOptions ppo = new PromptPointOptions($"\nSpécifiez le point d'insertion pour {compType} {currentDN} : ");
                        PromptPointResult ppr = ed.GetPoint(ppo);
                        if (ppr.Status != PromptStatus.OK) return;
                        insertPt = ppr.Value;

                        PromptPointOptions ppoDir = new PromptPointOptions("\nSpécifiez la direction : ");
                        ppoDir.UseBasePoint = true;
                        ppoDir.BasePoint = insertPt;
                        PromptPointResult pprDir = ed.GetPoint(ppoDir);
                        if (pprDir.Status != PromptStatus.OK) return;
                        direction = (pprDir.Value - insertPt).GetNormal();
                    }
                    else
                    {
                        return; // Annulé
                    }

                if (isOnPipe && selectedPolyId != ObjectId.Null)
                {
                    Curve curve = tr.GetObject(selectedPolyId, OpenMode.ForRead) as Curve;
                    if (curve != null)
                    {
                        double param = curve.GetParameterAtPoint(insertPt);
                        currentDN = PipeManager.GetActiveDNAtParameter(polyEnt, lineNumber, currentDN, param, tr);
                        if (LogiK3D.UI.MainPaletteControl.AvailableDiameters != null && LogiK3D.UI.MainPaletteControl.AvailableDiameters.ContainsKey(currentDN))
                        {
                            currentOD = LogiK3D.UI.MainPaletteControl.AvailableDiameters[currentDN];
                        }
                    }
                }

                string targetDN = currentDN;
                double targetOD = currentOD;

                // 2. Si c'est une réduction, demander le DN cible
                if (compType == "REDUCER" || compType == "RED_CONC" || compType == "RED_EXC")
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

                int pnVal = 16; // Valeur par défaut
                if (compType == "VANNE" || compType == "VALVE" || compType == "VALVE_GLOBE" || compType == "VALVE_BALL" || compType == "CHECK_VALVE" || 
                    compType == "FILTRE" || compType == "FILTER" || compType == "DEBIMETRE" || compType == "FLOWMETER" || 
                    compType == "DIAPHRAGME" || compType == "DIAPHRAGM" || compType == "BRIDE" || compType == "FLANGE")
                {
                    PromptKeywordOptions pkoPN = new PromptKeywordOptions("\nPression nominale [PN10/PN16/PN25/PN40] : ", "PN10 PN16 PN25 PN40");
                    pkoPN.AllowNone = true;
                    PromptResult prPN = ed.GetKeywords(pkoPN);
                    if (prPN.Status == PromptStatus.OK && !string.IsNullOrEmpty(prPN.StringResult))
                    {
                        int.TryParse(prPN.StringResult.Substring(2), out pnVal);
                    }
                }

                PipingGenerator generator = new PipingGenerator();
                ObjectId blockId = ObjectId.Null;
                string sapCode = "INCONNU";
                double compLength = 0;

                if (compType == "BRIDE" || compType == "FLANGE")
                {
                    int dnVal = 0;
                    if (currentDN.StartsWith("DN")) int.TryParse(currentDN.Substring(2), out dnVal);
                    else int.TryParse(currentDN, out dnVal);

                    var flangeData = LogiK3D.Specs.KohlerDatabase.FindFlange(currentDN);
                    
                    double flangeOD = InlineComponentDatabase.GetFlangeOuterDiameter(dnVal, pnVal);
                    double flangeThickness = flangeData != null && flangeData.WallThickness > 0 ? flangeData.WallThickness : 16.0;
                    double flangeLength = flangeThickness + 30.0;
                    compLength = flangeLength;

                    string blockName = $"FLANGE_{currentDN}_{flangeOD}_{flangeThickness}";
                    blockId = generator.GetOrCreateFlange(currentOD, flangeOD, flangeThickness, flangeLength, blockName);
                    sapCode = $"KOH-{currentDN}-FLANGE";

                    if (isOnPipe && selectedPolyId != ObjectId.Null)
                    {
                        Curve curve = tr.GetObject(selectedPolyId, OpenMode.ForRead) as Curve;
                        if (curve != null)
                        {
                            double param = curve.GetParameterAtPoint(insertPt);
                            if (param < curve.EndParam / 2.0)
                            {
                                direction = direction.Negate();
                            }
                        }
                    }
                    else
                    {
                        direction = direction.Negate();
                    }
                }
                else if (compType == "COUDE" || compType == "ELBOW_3D" || compType == "ELBOW_5D")
                {
                    double radius = compType == "ELBOW_5D" ? currentOD * 2.5 : currentOD * 1.5;
                    string blockName = $"ELBOW_90_{currentDN}_{radius}";
                    blockId = generator.GetOrCreateElbow(currentOD, radius, 90.0, blockName);
                    sapCode = $"KOH-{currentDN}-ELBOW";
                }
                else if (compType == "TEE")
                {
                    double length = currentOD * 2.0;
                    double branchHeight = currentOD;
                    string blockName = $"TEE_{currentDN}";
                    blockId = generator.GetOrCreateTee(currentOD, currentOD, length, branchHeight, blockName);
                    sapCode = $"KOH-{currentDN}-TEE";
                }
                else if (compType == "REDUCER" || compType == "RED_CONC" || compType == "RED_EXC")
                {
                    compLength = Math.Max(currentOD, targetOD) * 1.5;
                    string blockName = $"REDUCER_{Math.Max(currentOD, targetOD)}_{Math.Min(currentOD, targetOD)}";
                    blockId = generator.GetOrCreateReducer(Math.Max(currentOD, targetOD), Math.Min(currentOD, targetOD), compLength, blockName);
                    sapCode = $"KOH-{currentDN}-{targetDN}-RED";

                    if (currentOD < targetOD)
                    {
                        insertPt = insertPt + direction * compLength;
                        direction = direction.Negate();
                    }
                }
                else if (compType == "VANNE" || compType == "VALVE" || compType == "VALVE_GLOBE" || compType == "VALVE_BALL" || compType == "CHECK_VALVE")
                {
                    int dnVal = 0;
                    if (currentDN.StartsWith("DN")) int.TryParse(currentDN.Substring(2), out dnVal);
                    else int.TryParse(currentDN, out dnVal);

                    var dims = InlineComponentDatabase.GetDimensions(compType, dnVal, pnVal);
                    if (dims == null)
                    {
                        ed.WriteMessage($"\nErreur : DN {dnVal} non trouvé pour {compType}.");
                        return;
                    }
                    compLength = dims.Length;

                    PromptKeywordOptions pko = new PromptKeywordOptions("\nType d'actionneur [Manuel/Pneumatique] : ", "Manuel Pneumatique");
                    pko.AllowNone = true;
                    PromptResult pkr = ed.GetKeywords(pko);
                    bool isPneumatic = pkr.Status == PromptStatus.OK && pkr.StringResult == "Pneumatique";

                    string blockName = $"{compType}_DN{dnVal}_PN{pnVal}_{(isPneumatic ? "PNEU" : "MAN")}";
                    blockId = generator.GetOrCreateValve(dims.Length, dims.FlangeDiameter, dims.Height, isPneumatic, blockName);
                    sapCode = $"LOGIK-{currentDN}-{compType}";
                }
                else if (compType == "FILTRE" || compType == "FILTER")
                {
                    int dnVal = 0;
                    if (currentDN.StartsWith("DN")) int.TryParse(currentDN.Substring(2), out dnVal);
                    else int.TryParse(currentDN, out dnVal);

                    var dims = InlineComponentDatabase.GetDimensions(compType, dnVal, pnVal);
                    if (dims == null)
                    {
                        ed.WriteMessage($"\nErreur : DN {dnVal} non trouvé dans la base de données des filtres.");
                        return;
                    }
                    compLength = dims.Length;
                    string blockName = $"FILTER_Y_DN{dnVal}_PN{pnVal}";
                    blockId = generator.GetOrCreateYFilter(dims.Length, dims.FlangeDiameter, dims.Height, blockName);
                    sapCode = $"LOGIK-{currentDN}-FILTER";
                }
                else if (compType == "DEBIMETRE" || compType == "FLOWMETER")
                {
                    int dnVal = 0;
                    if (currentDN.StartsWith("DN")) int.TryParse(currentDN.Substring(2), out dnVal);
                    else int.TryParse(currentDN, out dnVal);

                    var dims = InlineComponentDatabase.GetDimensions(compType, dnVal, pnVal);
                    if (dims == null)
                    {
                        ed.WriteMessage($"\nErreur : DN {dnVal} non trouvé dans la base de données des débitmètres.");
                        return;
                    }
                    compLength = dims.Length;
                    string blockName = $"FLOWMETER_MAG_DN{dnVal}_PN{pnVal}";
                    blockId = generator.GetOrCreateFlowmeter(dims.Length, dims.FlangeDiameter, dims.Height, blockName);
                    sapCode = $"LOGIK-{currentDN}-FLOWMETER";
                }
                else if (compType == "DIAPHRAGME" || compType == "DIAPHRAGM")
                {
                    int dnVal = 0;
                    if (currentDN.StartsWith("DN")) int.TryParse(currentDN.Substring(2), out dnVal);
                    else int.TryParse(currentDN, out dnVal);

                    var dims = InlineComponentDatabase.GetDimensions(compType, dnVal, pnVal);
                    if (dims == null)
                    {
                        ed.WriteMessage($"\nErreur : DN {dnVal} non trouvé dans la base de données des diaphragmes.");
                        return;
                    }
                    compLength = dims.Length;
                    string blockName = $"DIAPHRAGM_DN{dnVal}_PN{pnVal}";
                    blockId = generator.GetOrCreateDiaphragm(dims.Length, dims.FlangeDiameter, dims.Height, blockName);
                    sapCode = $"LOGIK-{currentDN}-DIAPHRAGM";
                }

                // 5. Insérer le bloc et renseigner les attributs
                if (blockId != ObjectId.Null)
                {
                    BlockTableRecord btrSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    
                    // Si c'est une armature à brides, on insère aussi les brides
                    bool needsFlanges = (compType == "VANNE" || compType == "VALVE" || compType == "VALVE_GLOBE" || compType == "VALVE_BALL" || compType == "CHECK_VALVE" || 
                                         compType == "FILTRE" || compType == "FILTER" || compType == "DEBIMETRE" || compType == "FLOWMETER" || 
                                         compType == "DIAPHRAGME" || compType == "DIAPHRAGM");
                    
                    if (needsFlanges && isOnPipe)
                    {
                        int dnVal = 0;
                        if (currentDN.StartsWith("DN")) int.TryParse(currentDN.Substring(2), out dnVal);
                        else int.TryParse(currentDN, out dnVal);

                        var flangeData = LogiK3D.Specs.KohlerDatabase.FindFlange(currentDN);
                        double flangeOD = InlineComponentDatabase.GetFlangeOuterDiameter(dnVal, pnVal);
                        double flangeThickness = flangeData != null && flangeData.WallThickness > 0 ? flangeData.WallThickness : 16.0;
                        double flangeLength = flangeThickness + 30.0;

                        string flangeBlockName = $"FLANGE_{currentDN}_{flangeOD}_{flangeThickness}";
                        ObjectId flangeBlockId = generator.GetOrCreateFlange(currentOD, flangeOD, flangeThickness, flangeLength, flangeBlockName);
                        
                        if (flangeBlockId != ObjectId.Null)
                        {
                            // Bride 1 (avant)
                                Point3d f1Pt = insertPt - direction * flangeLength;
                                BlockReference f1 = new BlockReference(f1Pt, flangeBlockId);

                                Vector3d xAxis = Vector3d.XAxis;
                                double angle1 = xAxis.GetAngleTo(direction);
                                Vector3d axisOfRotation1 = xAxis.CrossProduct(direction);
                                if (!axisOfRotation1.IsZeroLength()) f1.TransformBy(Matrix3d.Rotation(angle1, axisOfRotation1, f1Pt));
                                else if (direction.X < 0) f1.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, f1Pt));

                                btrSpace.AppendEntity(f1);
                                tr.AddNewlyCreatedDBObject(f1, true);

                                // Attributs Bride 1
                                Dictionary<string, string> attrF1 = new Dictionary<string, string>
                                {
                                    { "POSITION", "" }, { "DESIGNATION", "BRIDE" }, { "DN", currentDN },
                                    { "PN", $"PN{pnVal}" }, { "ID", $"KOH-{currentDN}-FLANGE" }, { "NO LIGNE", lineNumber }
                                };
                                BlockTableRecord btrF1 = (BlockTableRecord)tr.GetObject(flangeBlockId, OpenMode.ForRead);
                                foreach (ObjectId entId in btrF1)
                                {
                                    if (tr.GetObject(entId, OpenMode.ForRead) is AttributeDefinition attrDef)
                                    {
                                        AttributeReference attrRef = new AttributeReference();
                                        attrRef.SetAttributeFromBlock(attrDef, f1.BlockTransform);
                                        if (attrF1.ContainsKey(attrDef.Tag.ToUpper())) attrRef.TextString = attrF1[attrDef.Tag.ToUpper()];
                                        f1.AttributeCollection.AppendAttribute(attrRef);
                                        tr.AddNewlyCreatedDBObject(attrRef, true);
                                    }
                                }

                                // Bride 2 (apr�s)
                                  Point3d f2Pt = insertPt + direction * (compLength + flangeLength);
                                  BlockReference f2 = new BlockReference(f2Pt, flangeBlockId);

                                  Vector3d dir2 = direction.Negate();
                                  double angle2 = xAxis.GetAngleTo(dir2);
                                  Vector3d axisOfRotation2 = xAxis.CrossProduct(dir2);
                                  if (!axisOfRotation2.IsZeroLength()) f2.TransformBy(Matrix3d.Rotation(angle2, axisOfRotation2, f2Pt));
                                  else if (dir2.X < 0) f2.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, f2Pt));

                                  btrSpace.AppendEntity(f2);
                                  tr.AddNewlyCreatedDBObject(f2, true);

                                  // Attributs Bride 2
                            foreach (ObjectId entId in btrF1)
                            {
                                if (tr.GetObject(entId, OpenMode.ForRead) is AttributeDefinition attrDef)
                                {
                                    AttributeReference attrRef = new AttributeReference();
                                    attrRef.SetAttributeFromBlock(attrDef, f2.BlockTransform);
                                    if (attrF1.ContainsKey(attrDef.Tag.ToUpper())) attrRef.TextString = attrF1[attrDef.Tag.ToUpper()];
                                    f2.AttributeCollection.AppendAttribute(attrRef);
                                    tr.AddNewlyCreatedDBObject(attrRef, true);
                                }
                            }
                        }
                    }

                    BlockReference bref = new BlockReference(insertPt, blockId);
                    
                    // Aligner le bloc avec la direction
                    Vector3d xAxisMain = Vector3d.XAxis;
                    double angleMain = xAxisMain.GetAngleTo(direction);
                    Vector3d axisOfRotationMain = xAxisMain.CrossProduct(direction);
                    
                    if (!axisOfRotationMain.IsZeroLength())
                    {
                        bref.TransformBy(Matrix3d.Rotation(angleMain, axisOfRotationMain, insertPt));
                    }
                    else if (direction.X < 0)
                    {
                        bref.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, insertPt));
                    }

                    btrSpace.AppendEntity(bref);
                    tr.AddNewlyCreatedDBObject(bref, true);

                    // Renseigner les attributs
                    Dictionary<string, string> attributes = new Dictionary<string, string>
                    {
                        { "POSITION", "" },
                        { "DESIGNATION", compType },
                        { "DN", currentDN },
                        { "PN", $"PN{pnVal}" },
                        { "ID", sapCode },
                        { "NO LIGNE", lineNumber }
                    };

                    if (compType == "REDUCER" || compType == "RED_CONC" || compType == "RED_EXC")
                    {
                        int currentDnVal = 0, targetDnVal = 0;
                        if (currentDN.StartsWith("DN")) int.TryParse(currentDN.Substring(2), out currentDnVal);
                        else int.TryParse(currentDN, out currentDnVal);
                        
                        if (targetDN.StartsWith("DN")) int.TryParse(targetDN.Substring(2), out targetDnVal);
                        else int.TryParse(targetDN, out targetDnVal);

                        if (currentDnVal >= targetDnVal)
                        {
                            attributes["GRAND DN"] = currentDN;
                            attributes["PETIT DN"] = targetDN;
                        }
                        else
                        {
                            attributes["GRAND DN"] = targetDN;
                            attributes["PETIT DN"] = currentDN;
                        }
                    }

                    BlockTableRecord btrBlock = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
                    foreach (ObjectId entId in btrBlock)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent is AttributeDefinition attrDef)
                        {
                            AttributeReference attrRef = new AttributeReference();
                            attrRef.SetAttributeFromBlock(attrDef, bref.BlockTransform);
                            
                            if (attributes.ContainsKey(attrDef.Tag.ToUpper()))
                            {
                                attrRef.TextString = attributes[attrDef.Tag.ToUpper()];
                            }
                            
                            bref.AttributeCollection.AppendAttribute(attrRef);
                            tr.AddNewlyCreatedDBObject(attrRef, true);
                        }
                    }
                    
                    ed.WriteMessage($"\n{compType} inséré avec succès.");

                    if (compType == "REDUCER" || compType == "RED_CONC")
                    {
                        if (LogiK3D.UI.MainPaletteControl.Instance != null)
                        {
                            LogiK3D.UI.MainPaletteControl.Instance.SetCurrentDN(targetDN);
                            ed.WriteMessage($"\n[Info] Le diamètre actif de la palette a été mis à jour sur {targetDN}.");
                        }
                    }
                    
                    // Mettre à jour la tuyauterie APRES avoir inséré le bloc
                    if (isOnPipe && polyEnt != null)
                    {
                        Curve pipeCurve = polyEnt as Curve;
                        if (pipeCurve != null)
                        {
                            try
                            {
                                PipeManager.DeleteLinkedSolids(pipeCurve, tr);
                                PipeManager.UpdateLineComponents(pipeCurve, lineNumber, currentDN, tr);
                                PipeManager.GeneratePiping(pipeCurve, lineNumber, currentDN, currentOD, currentThickness, tr);
                                ed.WriteMessage($"\nTuyauterie mise à jour.");
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nErreur lors de la mise à jour de la ligne : {ex.Message}");
                            }
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
            ed.WriteMessage("\n{0,-15} | {1,-20} | {2,-15} | {3,-20}", "Ligne", "Composant", "Longueur/DN", "Code SAP Kohler");
            ed.WriteMessage("\n-------------------------------------------------------------------------------");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                int count = 0;

                foreach (ObjectId id in btr)
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent != null)
                    {
                        // 1. Vérifier les solides (Tubes, Coudes) via XData
                        if (ent.XData != null)
                        {
                            ResultBuffer rb = ent.GetXDataForApplication(PipeManager.SolidDataAppName);
                            if (rb != null)
                            {
                                TypedValue[] values = rb.AsArray();
                                // Structure attendue : [0] RegApp, [1] GUID, [2] SAP Code, [3] Length, [4] Line Number, [5] CompType
                                if (values.Length >= 6)
                                {
                                    string sapCode = values[2].Value.ToString();
                                    double length = (double)values[3].Value;
                                    string lineNumber = values[4].Value.ToString();
                                    string compType = values[5].Value.ToString();

                                    ed.WriteMessage($"\n{lineNumber,-15} | {compType,-20} | {length,-15:F2} | {sapCode,-20}");
                                    count++;
                                    continue; // Passer à l'entité suivante
                                }
                            }
                        }

                        // 2. Vérifier les blocs (Brides, Tés, Réductions) via Attributs
                        if (ent is BlockReference bref)
                        {
                            string lineNumber = "INCONNU";
                            string compType = "INCONNU";
                            string sapCode = "INCONNU";
                            string dn = "";

                            foreach (ObjectId attId in bref.AttributeCollection)
                            {
                                AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                if (attRef != null)
                                {
                                    string tag = attRef.Tag.ToUpper();
                                    if (tag == "NO LIGNE") lineNumber = attRef.TextString;
                                    else if (tag == "DESIGNATION") compType = attRef.TextString;
                                    else if (tag == "ID") sapCode = attRef.TextString;
                                    else if (tag == "DN") dn = attRef.TextString;
                                }
                            }

                            if (compType != "INCONNU")
                            {
                                ed.WriteMessage($"\n{lineNumber,-15} | {compType,-20} | {dn,-15} | {sapCode,-20}");
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
        [CommandMethod("LOGIK_DRAW_VALVE")]
        public void DrawValve()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptIntegerOptions pio = new PromptIntegerOptions("\nEntrez le DN de la vanne (ex: 100) : ");
            PromptIntegerResult pir = ed.GetInteger(pio);
            if (pir.Status != PromptStatus.OK) return;
            int dn = pir.Value;

            var dims = InlineComponentDatabase.GetDimensions("VANNE", dn, 16);
            if (dims == null)
            {
                ed.WriteMessage($"\nErreur : DN {dn} non trouvé dans la base de données des vannes.");
                return;
            }

            PromptKeywordOptions pko = new PromptKeywordOptions("\nType d'actionneur [Manuel/Pneumatique] : ", "Manuel Pneumatique");
            PromptResult pkr = ed.GetKeywords(pko);
            if (pkr.Status != PromptStatus.OK) return;
            bool isPneumatic = pkr.StringResult == "Pneumatique";

            PromptPointOptions ppo = new PromptPointOptions("\nPoint d'insertion : ");
            PromptPointResult ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK) return;

            string blockName = $"VALVE_BUTTERFLY_DN{dn}_{(isPneumatic ? "PNEU" : "MAN")}";

            PipingGenerator generator = new PipingGenerator();
            ObjectId blockId = generator.GetOrCreateValve(dims.Length, dims.FlangeDiameter, dims.Height, isPneumatic, blockName);

            if (blockId != ObjectId.Null)
            {
                var attributes = new System.Collections.Generic.Dictionary<string, string>
                {
                    { "DN", dn.ToString() },
                    { "DESIGNATION", $"Vanne Papillon DN{dn} {(isPneumatic ? "Pneumatique" : "Manuelle")}" },
                    { "TYPE VANNE", "Papillon" },
                    { "ACTIONNEUR", isPneumatic ? "Pneumatique" : "Manuel" }
                };
                generator.InsertBlockReference(blockId, ppr.Value, attributes);
                ed.WriteMessage($"\nVanne insérée avec succès.");
            }
        }

        [CommandMethod("LOGIK_DRAW_FILTER")]
        public void DrawFilter()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptIntegerOptions pio = new PromptIntegerOptions("\nEntrez le DN du filtre (ex: 100) : ");
            PromptIntegerResult pir = ed.GetInteger(pio);
            if (pir.Status != PromptStatus.OK) return;
            int dn = pir.Value;

            var dims = InlineComponentDatabase.GetDimensions("FILTRE", dn, 16);
            if (dims == null)
            {
                ed.WriteMessage($"\nErreur : DN {dn} non trouvé dans la base de données des filtres.");
                return;
            }

            PromptPointOptions ppo = new PromptPointOptions("\nPoint d'insertion : ");
            PromptPointResult ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK) return;

            string blockName = $"FILTER_Y_DN{dn}";

            PipingGenerator generator = new PipingGenerator();
            ObjectId blockId = generator.GetOrCreateYFilter(dims.Length, dims.FlangeDiameter, dims.Height, blockName);

            if (blockId != ObjectId.Null)
            {
                var attributes = new System.Collections.Generic.Dictionary<string, string>
                {
                    { "DN", dn.ToString() },
                    { "DESIGNATION", $"Filtre Y DN{dn}" },
                    { "MAILLE", "Standard" }
                };
                generator.InsertBlockReference(blockId, ppr.Value, attributes);
                ed.WriteMessage($"\nFiltre inséré avec succès.");
            }
        }

        [CommandMethod("LOGIK_DRAW_FLOWMETER")]
        public void DrawFlowmeter()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptIntegerOptions pio = new PromptIntegerOptions("\nEntrez le DN du débitmètre (ex: 100) : ");
            PromptIntegerResult pir = ed.GetInteger(pio);
            if (pir.Status != PromptStatus.OK) return;
            int dn = pir.Value;

            var dims = InlineComponentDatabase.GetDimensions("DEBIMETRE", dn, 16);
            if (dims == null)
            {
                ed.WriteMessage($"\nErreur : DN {dn} non trouvé dans la base de données des débitmètres.");
                return;
            }

            PromptPointOptions ppo = new PromptPointOptions("\nPoint d'insertion : ");
            PromptPointResult ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK) return;

            string blockName = $"FLOWMETER_MAG_DN{dn}";

            PipingGenerator generator = new PipingGenerator();
            ObjectId blockId = generator.GetOrCreateFlowmeter(dims.Length, dims.FlangeDiameter, dims.Height, blockName);

            if (blockId != ObjectId.Null)
            {
                var attributes = new System.Collections.Generic.Dictionary<string, string>
                {
                    { "DN", dn.ToString() },
                    { "DESIGNATION", $"Débitmètre Électromagnétique DN{dn}" },
                    { "SIGNAL", "4-20mA" }
                };
                generator.InsertBlockReference(blockId, ppr.Value, attributes);
                ed.WriteMessage($"\nDébitmètre inséré avec succès.");
            }
        }

        [CommandMethod("LOGIK_DRAW_DIAPHRAGM")]
        public void DrawDiaphragm()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptIntegerOptions pio = new PromptIntegerOptions("\nEntrez le DN du diaphragme (ex: 100) : ");
            PromptIntegerResult pir = ed.GetInteger(pio);
            if (pir.Status != PromptStatus.OK) return;
            int dn = pir.Value;

            var dims = InlineComponentDatabase.GetDimensions("DIAPHRAGME", dn, 16);
            if (dims == null)
            {
                ed.WriteMessage($"\nErreur : DN {dn} non trouvé dans la base de données des diaphragmes.");
                return;
            }

            PromptPointOptions ppo = new PromptPointOptions("\nPoint d'insertion : ");
            PromptPointResult ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK) return;

            string blockName = $"DIAPHRAGM_DN{dn}";

            PipingGenerator generator = new PipingGenerator();
            ObjectId blockId = generator.GetOrCreateDiaphragm(dims.Length, dims.FlangeDiameter, dims.Height, blockName);

            if (blockId != ObjectId.Null)
            {
                var attributes = new System.Collections.Generic.Dictionary<string, string>
                {
                    { "DN", dn.ToString() },
                    { "DESIGNATION", $"Plaque à orifice (Diaphragme) DN{dn}" },
                    { "ORIFICE", "Calculé" }
                };
                generator.InsertBlockReference(blockId, ppr.Value, attributes);
                ed.WriteMessage($"\nDiaphragme inséré avec succès.");
            }
        }
    }
}


