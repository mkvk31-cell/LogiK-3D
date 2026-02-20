using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using HtmlAgilityPack;
using CsvHelper;
using System.Globalization;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace PdfReaderTest
{
    class WebScraper
    {
        public class FittingDimension
        {
            public string? Type { get; set; }
            public string? DN { get; set; }
            public string? OuterDiameter { get; set; }
            public string? WallThickness { get; set; }
            public string? Radius { get; set; }
            public string? Length { get; set; }
            public string? Material { get; set; }
            public string? PN { get; set; }
            public string? Url { get; set; }
            public string? ImageUrl { get; set; }
            public string? TechnicalDrawingUrl { get; set; }
        }

        public static async Task RunAsync()
        {
            var dimensions = new List<FittingDimension>();
            using var httpClient = new HttpClient();
            
            // Base URLs for the specific categories we want
            var categoryUrls = new List<string>
            {
                "https://kohler.ch/fr/categorie/coudes",
                "https://kohler.ch/fr/categorie/tes",
                "https://kohler.ch/fr/categorie/reductions-2",
                "https://kohler.ch/fr/categorie/collerettes",
                "https://kohler.ch/fr/categorie/brides-plates",
                "https://kohler.ch/fr/categorie/brides-a-visser",
                "https://kohler.ch/fr/categorie/brides-a-collerettes",
                "https://kohler.ch/fr/categorie/brides-libres-et-mobiles",
                "https://kohler.ch/fr/categorie/brides-pleines"
            };
            
            var productLinks = new List<string>();
            
            foreach (var baseUrl in categoryUrls)
            {
                Console.WriteLine($"Fetching products from {baseUrl}...");
                
                try
                {
                    var html = await httpClient.GetStringAsync(baseUrl);
                    var doc = new HtmlDocument();
                    doc.LoadHtml(html);
                    
                    var scriptTags = doc.DocumentNode.SelectNodes("//script[@type='application/json' and @data-sveltekit-fetched]");
                    
                    if (scriptTags != null)
                    {
                        foreach (var script in scriptTags)
                        {
                            try
                            {
                                var json = JObject.Parse(script.InnerText);
                                var bodyStr = json["body"]?.ToString();
                                if (!string.IsNullOrEmpty(bodyStr))
                                {
                                    var bodyJson = JObject.Parse(bodyStr);
                                    var slugs = bodyJson.SelectTokens("$..products.items[*].slug").Select(t => t.ToString()).Where(s => s.Contains("-")).Distinct().ToList();
                                    foreach (var slug in slugs)
                                    {
                                        if (!productLinks.Contains($"https://kohler.ch/fr/produit/{slug}"))
                                        {
                                            productLinks.Add($"https://kohler.ch/fr/produit/{slug}");
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to fetch category {baseUrl}: {ex.Message}");
                }
            }
                
            // If we didn't find any, let's just hardcode the main ones we know
            if (productLinks.Count == 0)
            {
                productLinks.Add("https://kohler.ch/fr/produit/103010r-1651"); // Elbows
                productLinks.Add("https://kohler.ch/fr/produit/103025r-166");  // Tees
                productLinks.Add("https://kohler.ch/fr/produit/103030r-160");  // Reducers
                productLinks.Add("https://kohler.ch/fr/produit/103040r-157");  // Flanges
            }
                                 
            Console.WriteLine($"Found {productLinks.Count} product links.");
            
            foreach (var link in productLinks)
                {
                    string fullUrl = link.StartsWith("http") ? link : $"https://kohler.ch{link}";
                    Console.WriteLine($"Processing {fullUrl}...");
                    
                    try
                    {
                        var productHtml = await httpClient.GetStringAsync(fullUrl);
                        
                        var productDoc = new HtmlDocument();
                        productDoc.LoadHtml(productHtml);
                        
                        var productScriptTags = productDoc.DocumentNode.SelectNodes("//script[@type='application/json' and @data-sveltekit-fetched]");
                        var variants = new List<JToken>();
                        
                        if (productScriptTags != null)
                        {
                            foreach (var script in productScriptTags)
                            {
                                try
                                {
                                    var json = JObject.Parse(script.InnerText);
                                    var bodyStr = json["body"]?.ToString();
                                    if (!string.IsNullOrEmpty(bodyStr))
                                    {
                                        var bodyJson = JObject.Parse(bodyStr);
                                        var foundVariants = bodyJson.SelectTokens("$..productVariantSearch.items[*]");
                                        if (foundVariants.Any())
                                        {
                                            variants.AddRange(foundVariants);
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                            
                            Console.WriteLine($"Found {variants.Count} variants.");
                            
                            foreach (var variant in variants)
                            {
                                var options = variant["options"] as JArray;
                                if (options != null)
                                {
                                    string d = "", s = "", r = "", l = "", type = "", mat = "", pn = "", imageUrl = "", techDrawingUrl = "", dn = "";
                                    
                                    // Get the product name from the parent or variant
                                    var productNode = variant["product"];
                                    if (productNode != null)
                                    {
                                        type = productNode["name"]?.ToString() ?? "";
                                        
                                        // Get images
                                        var assets = productNode["assets"] as JArray;
                                        if (assets != null)
                                        {
                                            foreach (var asset in assets)
                                            {
                                                var tags = asset["tags"] as JArray;
                                                if (tags != null)
                                                {
                                                    bool isProductImage = tags.Any(t => t["value"]?.ToString() == "product-image");
                                                    bool isTechDrawing = tags.Any(t => t["value"]?.ToString() == "technical-drawing");
                                                    
                                                    if (isProductImage) imageUrl = asset["source"]?.ToString() ?? "";
                                                    if (isTechDrawing) techDrawingUrl = asset["source"]?.ToString() ?? "";
                                                }
                                            }
                                        }
                                    }
                                    
                                    foreach (var option in options)
                                    {
                                        var groupName = option["group"]?["name"]?.ToString();
                                        var value = option["name"]?.ToString();
                                        
                                        if (groupName == "Ø extérieur mm") d = value;
                                        else if (groupName == "Épaisseur de paroi mm" || groupName == "Épaisseur mm") s = value;
                                        else if (groupName == "Rayon mm") r = value;
                                        else if (groupName == "Longueur totale mm") l = value;
                                        else if (groupName == "Matériau") mat = value;
                                        else if (groupName == "PN") pn = value;
                                        else if (groupName == "pour tube mm / trou central mm" || groupName == "pour tube mm") 
                                        {
                                            // Extract the pipe diameter (first part before /)
                                            var parts = value.Split('/');
                                            if (parts.Length > 0)
                                            {
                                                dn = parts[0].Trim();
                                            }
                                        }
                                    }
                                    
                                    if (!string.IsNullOrEmpty(d))
                                    {
                                        dimensions.Add(new FittingDimension
                                        {
                                            Type = type,
                                            DN = dn,
                                            OuterDiameter = d,
                                            WallThickness = s,
                                            Radius = r,
                                            Length = l,
                                            Material = mat,
                                            PN = pn,
                                            Url = fullUrl,
                                            ImageUrl = imageUrl,
                                            TechnicalDrawingUrl = techDrawingUrl
                                        });
                                    }
                                }
                            }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to process {fullUrl}: {ex.Message}");
                    }
                }
                
                // Export to CSV
                string csvPath = "web_fittings_dimensions.csv";
                using (var writer = new StreamWriter(csvPath))
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteRecords(dimensions);
                }
                Console.WriteLine($"\nExported all {dimensions.Count} dimensions to {csvPath}");
        }
    }
}