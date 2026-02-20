using System;
using System.Collections.Generic;
using Autodesk.ProcessPower.PartsRepository.Specification;
using Autodesk.ProcessPower.DataObjects;

namespace LogiK3D.Piping
{
    public class SpecReader
    {
        public static List<string> GetAvailableSpecs(string specFolderPath)
        {
            List<string> specs = new List<string>();
            try
            {
                if (System.IO.Directory.Exists(specFolderPath))
                {
                    string[] files = System.IO.Directory.GetFiles(specFolderPath, "*.pspx");
                    foreach (string file in files)
                    {
                        specs.Add(System.IO.Path.GetFileNameWithoutExtension(file));
                    }
                }
            }
            catch (Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nErreur lecture specs: {ex.Message}");
            }
            return specs;
        }

        public static double GetPipeOuterDiameter(string specPath, string dn)
        {
            // Valeur par défaut si non trouvée
            double od = 114.3; 
            
            try
            {
                // Utilisation de l'API Plant 3D pour lire la spec sans créer de projet
                PipePartSpecification spec = PipePartSpecification.OpenSpecification(specPath);
                
                if (spec != null)
                {
                    PnPDatabase db = spec.Database;
                    if (db != null)
                    {
                        PnPTable table = db.Tables["EngineeringItems"];
                        if (table != null)
                        {
                            // On cherche un tuyau (Pipe) avec le bon DN
                            string query = $"PartCategory='Pipe' AND NominalDiameter='{dn}'";
                            PnPRow[] rows = table.Select(query);
                            
                            if (rows != null && rows.Length > 0)
                            {
                                // On prend le premier résultat
                                PnPRow row = rows[0];
                                
                                // Le diamètre extérieur est souvent dans MatchingPipeOd ou OutsideDiameter
                                try
                                {
                                    if (row["MatchingPipeOd"] != DBNull.Value)
                                    {
                                        od = Convert.ToDouble(row["MatchingPipeOd"]);
                                    }
                                }
                                catch
                                {
                                    try
                                    {
                                        if (row["OutsideDiameter"] != DBNull.Value)
                                        {
                                            od = Convert.ToDouble(row["OutsideDiameter"]);
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                    spec.Close();
                }
            }
            catch (Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nErreur lecture OD: {ex.Message}");
            }
            
            return od;
        }
    }
}