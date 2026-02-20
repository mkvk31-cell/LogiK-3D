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
                            if (ent != null)
                            {
                                if (ent.XData != null)
                                {
                                    ResultBuffer rb = ent.GetXDataForApplication("LogiK_Data");
                                    if (rb != null)
                                    {
                                        TypedValue[] tvs = rb.AsArray();
                                        if (tvs.Length >= 10)
                                        {
                                            string lineNumber = tvs[4].Value.ToString();
                                            if (!lineGroups.ContainsKey(lineNumber))
                                                lineGroups[lineNumber] = new System.Collections.Generic.List<Entity>();
                                            lineGroups[lineNumber].Add(ent);
                                        }
                                    }
                                }
                                else if (ent is BlockReference bref)
                                {
                                    string lineNumber = "INCONNU";
                                    foreach (ObjectId attId in bref.AttributeCollection)
                                    {
                                        AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                        if (attRef != null && attRef.Tag.ToUpper() == "NO LIGNE")
                                        {
                                            lineNumber = attRef.TextString;
                                            break;
                                        }
                                    }
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
                                if (ent.XData != null)
                                {
                                    ResultBuffer rb = ent.GetXDataForApplication("LogiK_Data");
                                    if (rb != null)
                                    {
                                        TypedValue[] tvs = rb.AsArray();
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
                                    }
                                }
                                else if (ent is BlockReference bref)
                                {
                                    string compType = "INCONNU";
                                    string sapCode = "INCONNU";
                                    string dn = "";
                                    
                                    foreach (ObjectId attId in bref.AttributeCollection)
                                    {
                                        AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                        if (attRef != null)
                                        {
                                            string tag = attRef.Tag.ToUpper();
                                            if (tag == "DESIGNATION") compType = attRef.TextString;
                                            else if (tag == "ID") sapCode = attRef.TextString;
                                            else if (tag == "DN") dn = attRef.TextString;
                                        }
                                    }
                                    
                                    double dnValue = 0;
                                    if (dn.StartsWith("DN")) double.TryParse(dn.Substring(2), out dnValue);
                                    else double.TryParse(dn, out dnValue);
                                    
                                    // Récupérer les points de connexion (ports) depuis la définition du bloc
                                    List<Point3d> ports = new List<Point3d>();
                                    BlockTableRecord btrBlock = tr.GetObject(bref.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                    if (btrBlock != null)
                                    {
                                        foreach (ObjectId entId in btrBlock)
                                        {
                                            Entity bEnt = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                            if (bEnt is DBPoint pt)
                                            {
                                                ports.Add(pt.Position.TransformBy(bref.BlockTransform));
                                            }
                                        }
                                    }
                                    
                                    // Assigner les points (fallback sur la position du bloc si non trouvés)
                                    Point3d p1 = ports.Count > 0 ? ports[0] : bref.Position;
                                    Point3d p2 = ports.Count > 1 ? ports[1] : bref.Position;
                                    Point3d p3 = ports.Count > 2 ? ports[2] : bref.Position;
                                    Point3d center = bref.Position;
                                    
                                    if (compType.Contains("COUDE") || compType.Contains("ELBOW"))
                                    {
                                        writer.WriteLine("ELBOW");
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p1.X, p1.Y, p1.Z, dnValue));
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p2.X, p2.Y, p2.Z, dnValue));
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    CENTRE-POINT {0:F4} {1:F4} {2:F4}", center.X, center.Y, center.Z));
                                        writer.WriteLine($"    ITEM-CODE {sapCode}");
                                        writer.WriteLine($"    SKEY ELBW");
                                        writer.WriteLine($"    FABRICATION-ITEM");
                                    }
                                    else if (compType.Contains("BRIDE") || compType.Contains("FLANGE"))
                                    {
                                        writer.WriteLine("FLANGE");
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p1.X, p1.Y, p1.Z, dnValue));
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p2.X, p2.Y, p2.Z, dnValue));
                                        writer.WriteLine($"    ITEM-CODE {sapCode}");
                                        writer.WriteLine($"    SKEY FLWN");
                                        writer.WriteLine($"    FABRICATION-ITEM");
                                    }
                                    else if (compType.Contains("TEE"))
                                    {
                                        writer.WriteLine("TEE");
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p1.X, p1.Y, p1.Z, dnValue));
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p2.X, p2.Y, p2.Z, dnValue));
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    CENTRE-POINT {0:F4} {1:F4} {2:F4}", center.X, center.Y, center.Z));
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    BRANCH1-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p3.X, p3.Y, p3.Z, dnValue));
                                        writer.WriteLine($"    ITEM-CODE {sapCode}");
                                        writer.WriteLine($"    SKEY TEBW");
                                        writer.WriteLine($"    FABRICATION-ITEM");
                                    }
                                    else if (compType.Contains("REDUCER") || compType.Contains("RED_CONC") || compType.Contains("RED_EXC"))
                                    {
                                        writer.WriteLine("REDUCER-CONCENTRIC");
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p1.X, p1.Y, p1.Z, dnValue));
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p2.X, p2.Y, p2.Z, dnValue));
                                        writer.WriteLine($"    ITEM-CODE {sapCode}");
                                        writer.WriteLine($"    SKEY RCON");
                                        writer.WriteLine($"    FABRICATION-ITEM");
                                    }
                                    else if (compType.Contains("VANNE") || compType.Contains("VALVE"))
                                    {
                                        writer.WriteLine("VALVE");
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p1.X, p1.Y, p1.Z, dnValue));
                                        writer.WriteLine(string.Format(System.Globalization.CultureInfo.InvariantCulture, "    END-POINT {0:F4} {1:F4} {2:F4} {3:0.####}", p2.X, p2.Y, p2.Z, dnValue));
                                        writer.WriteLine($"    ITEM-CODE {sapCode}");
                                        writer.WriteLine($"    SKEY VVFL");
                                        writer.WriteLine($"    FABRICATION-ITEM");
                                    }
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