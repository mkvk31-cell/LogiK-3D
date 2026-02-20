using System;
using System.IO;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace LogiK3D.Piping
{
    public class Plant3DCommands
    {
        [CommandMethod("LOGIK_EXPORT_PCF")]
        public void ExportPCF()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 1. Demander à l'utilisateur de sélectionner les composants LogiK 3D
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nSélectionnez les conduites et composants pour générer le PCF : ";
            
            PromptSelectionResult psr = ed.GetSelection(pso);

            if (psr.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nSélection annulée.");
                return;
            }

            // 2. Générer notre propre fichier PCF sans utiliser Plant 3D
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string pcfPath = Path.Combine(desktopPath, $"LogiK3D_Export_{DateTime.Now:yyyyMMdd_HHmmss}.pcf");

                using (StreamWriter writer = new StreamWriter(pcfPath, false, Encoding.ASCII))
                {
                    // En-tête standard PCF
                    writer.WriteLine("ISOMETRIC-DEF-FILE");
                    writer.WriteLine("UNITS-BORE MM");
                    writer.WriteLine("UNITS-CO-ORDS MM");
                    writer.WriteLine("UNITS-WEIGHT KGS");
                    writer.WriteLine("");

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        foreach (SelectedObject so in psr.Value)
                        {
                            Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                            if (ent is Solid3d solid)
                            {
                                // Extraction des données XData
                                string sapCode = "UNKNOWN";
                                double length = 0;
                                
                                if (ent.XData != null)
                                {
                                    TypedValue[] tvs = ent.XData.AsArray();
                                    foreach (TypedValue tv in tvs)
                                    {
                                        if (tv.TypeCode == (int)DxfCode.ExtendedDataAsciiString && tv.Value.ToString().StartsWith("KOH-"))
                                            sapCode = tv.Value.ToString();
                                        if (tv.TypeCode == (int)DxfCode.ExtendedDataReal)
                                            length = (double)tv.Value;
                                    }
                                }

                                // Approximation des points de départ et d'arrivée basée sur la BoundingBox
                                // (Dans une version avancée, on stockerait les points exacts dans les XData)
                                Point3d min = solid.GeometricExtents.MinPoint;
                                Point3d max = solid.GeometricExtents.MaxPoint;

                                writer.WriteLine("PIPE");
                                writer.WriteLine($"    END-POINT {min.X:F2} {min.Y:F2} {min.Z:F2} {LogiK3D.UI.MainPaletteControl.CurrentOuterDiameter}");
                                writer.WriteLine($"    END-POINT {max.X:F2} {max.Y:F2} {max.Z:F2} {LogiK3D.UI.MainPaletteControl.CurrentOuterDiameter}");
                                writer.WriteLine($"    ITEM-CODE {sapCode}");
                                writer.WriteLine($"    FABRICATION-ITEM");
                            }
                        }
                        tr.Commit();
                    }
                }

                ed.WriteMessage($"\nFichier PCF généré avec succès sur le Bureau : {pcfPath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nErreur lors de l'export PCF : {ex.Message}");
            }
        }
    }
}