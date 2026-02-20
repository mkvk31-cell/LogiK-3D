using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace LogiK3D.Specs
{
    public class KohlerComponent
    {
        public string Type { get; set; }
        public string DN { get; set; }
        public double OuterDiameter { get; set; }
        public double WallThickness { get; set; }
        public double Radius { get; set; }
        public double Length { get; set; }
        public string Material { get; set; }
        public string PN { get; set; }
        public string Url { get; set; }
        public string ImageUrl { get; set; }
        public string TechnicalDrawingUrl { get; set; }
    }

    public static class KohlerDatabase
    {
        private static List<KohlerComponent> _components = new List<KohlerComponent>();
        private static bool _isLoaded = false;

        public static void LoadDatabase()
        {
            if (_isLoaded) return;

            try
            {
                // Find the CSV file. It should be in the same directory as the DLL or in a known location.
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string directory = Path.GetDirectoryName(assemblyPath);
                
                // For development, we might need to look in the PdfReaderTest folder
                string csvPath = Path.Combine(directory, "web_fittings_dimensions.csv");
                
                if (!File.Exists(csvPath))
                {
                    // Try to find it in the project directory (useful during dev)
                    string projectDir = Path.GetFullPath(Path.Combine(directory, @"..\..\..\PdfReaderTest"));
                    csvPath = Path.Combine(projectDir, "web_fittings_dimensions.csv");
                }

                if (File.Exists(csvPath))
                {
                    var lines = File.ReadAllLines(csvPath).Skip(1); // Skip header
                    foreach (var line in lines)
                    {
                        // Simple CSV parsing (doesn't handle commas inside quotes perfectly, but good enough for our data)
                        // A better approach would be to use a CSV parser library, but we'll do a basic split for now
                        // Since the Type column has quotes and commas, we need a slightly smarter split
                        
                        var parts = ParseCsvLine(line);
                        if (parts.Count >= 11)
                        {
                            var comp = new KohlerComponent
                            {
                                Type = parts[0].Trim('"'),
                                DN = parts[1],
                                Material = parts[6],
                                PN = parts[7],
                                Url = parts[8],
                                ImageUrl = parts[9],
                                TechnicalDrawingUrl = parts[10]
                            };

                            if (double.TryParse(parts[2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double od))
                                comp.OuterDiameter = od;
                                
                            if (double.TryParse(parts[3], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double wt))
                                comp.WallThickness = wt;
                                
                            if (double.TryParse(parts[4], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double r))
                                comp.Radius = r;
                                
                            if (double.TryParse(parts[5], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double l))
                                comp.Length = l;

                            _components.Add(comp);
                        }
                    }
                    _isLoaded = true;
                }
            }
            catch (Exception ex)
            {
                // Log error
                System.Diagnostics.Debug.WriteLine($"Error loading Kohler database: {ex.Message}");
            }
        }

        private static List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            int start = 0;

            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (line[i] == ',' && !inQuotes)
                {
                    result.Add(line.Substring(start, i - start));
                    start = i + 1;
                }
            }
            result.Add(line.Substring(start));
            return result;
        }

        public static KohlerComponent FindFlange(string dn, string pn = "")
        {
            LoadDatabase();
            
            var query = _components.Where(c => c.Type.Contains("Bride", StringComparison.OrdinalIgnoreCase));
            
            // Clean up DN (e.g., "DN100" -> "100")
            string cleanDn = dn.Replace("DN", "").Trim();
            
            // Try to match by DN first
            var dnMatches = query.Where(c => c.DN == cleanDn || c.DN.StartsWith(cleanDn + ".")).ToList();
            
            if (dnMatches.Any())
            {
                if (!string.IsNullOrEmpty(pn))
                {
                    var pnMatches = dnMatches.Where(c => c.Type.Contains(pn, StringComparison.OrdinalIgnoreCase)).ToList();
                    if (pnMatches.Any()) return pnMatches.First();
                }
                return dnMatches.First();
            }
            
            // Fallback: try to match by OuterDiameter if DN didn't work
            // This is a bit tricky since flanges have different ODs than pipes, but we can try
            
            return query.FirstOrDefault(); // Just return the first one if no match
        }
    }
}