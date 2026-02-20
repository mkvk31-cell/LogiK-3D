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
                        // Grouper les composants par numéro de ligne
                        var lineGroups = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<Entity>>();

                        foreach (SelectedObject so in psr.Value)
                        {
                            Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                            if (ent is Solid3d solid && ent.XData != null)
                            {
                                TypedValue[] tvs = ent.XData.AsArray();
                                if (tvs.Length >= 10 && tvs[0].Value.ToString() == "LogiK_Data")
                                {
                                    string lineNumber = tvs[4].Value.ToString();
                                    if (!lineGroups.ContainsKey(lineNumber))
                                        lineGroups[lineNumber] = new System.Collections.Generic.List<Entity>();
                                    lineGroups[lineNumber].Add(ent);
                                }
                            }
                        }

                        foreach (var group in lineGroups)
                        {
                            writer.WriteLine($"PIPELINE-REFERENCE {group.Key}");
                            writer.WriteLine("");

                            foreach (Entity ent in group.Value)
                            {
                                TypedValue[] tvs = ent.XData.AsArray();
                                string sapCode = tvs[2].Value.ToString();
                                double length = (double)tvs[3].Value;
                                string compType = tvs[5].Value.ToString();
                                double dnValue = (double)tvs[6].Value;
                                Point3d p1 = (Point3d)tvs[7].Value;
                                Point3d p2 = (Point3d)tvs[8].Value;
                                Point3d p3 = (Point3d)tvs[9].Value;

                                if (compType == "PIPE")
                                {
                                    writer.WriteLine("PIPE");
                                    writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p1.X, p1.Y, p1.Z, dnValue));
                                    writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p2.X, p2.Y, p2.Z, dnValue));
                                    writer.WriteLine($"    ITEM-CODE {sapCode}");
                                    writer.WriteLine($"    SKEY PBFL");
                                    writer.WriteLine($"    FABRICATION-ITEM");
                                }
                                else if (compType == "ELBOW")
                                {
                                    writer.WriteLine("ELBOW");
                                    writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p1.X, p1.Y, p1.Z, dnValue));
                                    writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p2.X, p2.Y, p2.Z, dnValue));
                                    writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    CENTRE-POINT {0:F4} {1:F4} {2:F4}", p3.X, p3.Y, p3.Z));
                                    writer.WriteLine($"    ITEM-CODE {sapCode}");
                                    writer.WriteLine($"    SKEY ELBW");
                                    writer.WriteLine($"    FABRICATION-ITEM");
                                }
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