using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace LogiK3D.Specs
{
    public static class PdfSpecParser
    {
        // Dictionnaire des diamètres extérieurs standards (EN 10220 / ISO 1127)
        private static readonly Dictionary<string, double> StandardODs = new Dictionary<string, double>
        {
            { "15", 21.3 },
            { "20", 26.9 },
            { "25", 33.7 },
            { "32", 42.4 },
            { "40", 48.3 },
            { "50", 60.3 },
            { "65", 76.1 },
            { "80", 88.9 },
            { "100", 114.3 },
            { "125", 139.7 },
            { "150", 168.3 },
            { "200", 219.1 },
            { "250", 273.0 },
            { "300", 323.9 }
        };

        public static List<PipingSpec> ParsePdf(string pdfPath)
        {
            var specs = new List<PipingSpec>();

            try
            {
                using (PdfDocument document = PdfDocument.Open(pdfPath))
                {
                    PipingSpec currentSpec = null;

                    foreach (Page page in document.GetPages())
                    {
                        string text = page.Text;
                        
                        // Chercher le début d'une classe, ex: "2.3. Classe B031 : Acier carbone PN16"
                        var classMatch = Regex.Match(text, @"Classe\s+([A-Z0-9]{4})\s*:\s*(.*?)(?:\s*-\s*type|\r|\n|$)");
                        if (classMatch.Success)
                        {
                            currentSpec = new PipingSpec();
                            currentSpec.Name = classMatch.Groups[1].Value;
                            currentSpec.Description = classMatch.Groups[2].Value.Trim();
                            currentSpec.Material = classMatch.Groups[2].Value.Trim();
                            currentSpec.Components = new List<SpecComponent>();
                            specs.Add(currentSpec);
                        }

                        if (currentSpec != null)
                        {
                            // Extraction des plages de DN et des épaisseurs
                            var dnMatch = Regex.Match(text, @"Tubes\s+([\d\s\-]+)\s+Tube");
                            var tubeMatches = Regex.Matches(text, @"(?i)Tube.*?ép\.\s*(\d+[\.,]\d+)");
                            
                            if (dnMatch.Success && tubeMatches.Count > 0)
                            {
                                string dnText = dnMatch.Groups[1].Value;
                                var tokens = dnText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                
                                int half = tokens.Length / 2;
                                List<string> starts = new List<string>();
                                List<string> ends = new List<string>();
                                
                                for (int i = 0; i < half; i++)
                                {
                                    starts.Add(tokens[i]);
                                    ends.Add(tokens[i + half]);
                                }
                                
                                int count = Math.Min(starts.Count, tubeMatches.Count);
                                
                                for (int i = 0; i < count; i++)
                                {
                                    string startDnStr = starts[i];
                                    string endDnStr = ends[i];
                                    if (!double.TryParse(tubeMatches[i].Groups[1].Value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double thickness))
                                    {
                                        continue;
                                    }
                                    
                                    List<string> dnsToAdd = new List<string>();
                                    
                                    if (startDnStr != "-")
                                    {
                                        dnsToAdd.Add(startDnStr);
                                        
                                        if (endDnStr != "-" && int.TryParse(startDnStr, out int startDn) && int.TryParse(endDnStr, out int endDn))
                                        {
                                            // Ajouter les DN intermédiaires standards
                                            foreach (var stdDn in StandardODs.Keys)
                                            {
                                                if (int.TryParse(stdDn, out int stdDnInt) && stdDnInt > startDn && stdDnInt <= endDn)
                                                {
                                                    dnsToAdd.Add(stdDn);
                                                }
                                            }
                                        }
                                    }
                                    
                                    foreach (var dn in dnsToAdd)
                                    {
                                        double od = StandardODs.ContainsKey(dn) ? StandardODs[dn] : 0;
                                        
                                        // Ajouter le tube
                                        currentSpec.Components.Add(new SpecComponent
                                        {
                                            Type = "PIPE",
                                            DN = dn,
                                            OuterDiameter = od,
                                            Thickness = thickness,
                                            Schedule = "STD"
                                        });

                                        // Ajouter le coude 90 (3D par défaut)
                                        currentSpec.Components.Add(new SpecComponent
                                        {
                                            Type = "ELBOW",
                                            DN = dn,
                                            OuterDiameter = od,
                                            Thickness = thickness,
                                            Schedule = "3D"
                                        });

                                        // Ajouter le Té
                                        currentSpec.Components.Add(new SpecComponent
                                        {
                                            Type = "TEE",
                                            DN = dn,
                                            OuterDiameter = od,
                                            Thickness = thickness,
                                            Schedule = "STD"
                                        });
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log l'erreur au lieu d'afficher une MessageBox qui peut faire crasher AutoCAD
                System.Diagnostics.Debug.WriteLine($"Erreur lors de la lecture du PDF : {ex.Message}");
            }

            // Si aucune classe n'a été trouvée (format différent), on crée une spec générique
            if (specs.Count == 0)
            {
                var fallbackSpec = new PipingSpec();
                fallbackSpec.Name = Path.GetFileNameWithoutExtension(pdfPath);
                fallbackSpec.Description = "Importé depuis PDF (Format non reconnu)";
                fallbackSpec.Material = "Inconnu";
                fallbackSpec.Components = new List<SpecComponent>();
                specs.Add(fallbackSpec);
            }

            return specs;
        }
    }
}
