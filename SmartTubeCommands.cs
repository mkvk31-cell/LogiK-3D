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