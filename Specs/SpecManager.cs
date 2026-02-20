using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace LogiK3D.Specs
{
    public static class SpecManager
    {
        private static string GetSpecFilePath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "LogiK3D");
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return Path.Combine(folder, "specs.json");
        }

        public static List<PipingSpec> LoadSpecs()
        {
            string path = GetSpecFilePath();
            if (!File.Exists(path))
            {
                return CreateDefaultSpecs();
            }

            try
            {
                string json = File.ReadAllText(path);
                var specs = JsonConvert.DeserializeObject<List<PipingSpec>>(json);
                return specs ?? new List<PipingSpec>();
            }
            catch
            {
                return CreateDefaultSpecs();
            }
        }

        public static void SaveSpecs(List<PipingSpec> specs)
        {
            string path = GetSpecFilePath();
            string json = JsonConvert.SerializeObject(specs, Formatting.Indented);
            File.WriteAllText(path, json);
        }

        private static List<PipingSpec> CreateDefaultSpecs()
        {
            var specs = new List<PipingSpec>
            {
                new PipingSpec 
                { 
                    Name = "CS150", 
                    Description = "Acier Carbone 150#", 
                    Material = "Carbon Steel",
                    Components = new List<SpecComponent>
                    {
                        new SpecComponent { Type = "PIPE", DN = "100", OuterDiameter = 114.3, Thickness = 6.02, Schedule = "Sch 40", SAPCode = "PIP-CS-100-40" },
                        new SpecComponent { Type = "ELBOW", DN = "100", OuterDiameter = 114.3, Thickness = 6.02, Schedule = "Sch 40", SAPCode = "ELB-CS-100-40" }
                    }
                },
                new PipingSpec 
                { 
                    Name = "SS150", 
                    Description = "Inox 150#", 
                    Material = "Stainless Steel",
                    Components = new List<SpecComponent>
                    {
                        new SpecComponent { Type = "PIPE", DN = "100", OuterDiameter = 114.3, Thickness = 3.05, Schedule = "Sch 10S", SAPCode = "PIP-SS-100-10S" }
                    }
                }
            };
            
            SaveSpecs(specs);
            return specs;
        }
    }
}
