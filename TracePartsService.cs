using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;

namespace LogiK3D.Integration
{
    /// <summary>
    /// Service pour interagir avec l'API officielle de TraceParts.
    /// Nécessite une clé API valide depuis https://developers.traceparts.com/
    /// </summary>
    public class TracePartsService
    {
        private readonly HttpClient _httpClient;
        private const string ApiBaseUrl = "https://gateway.traceparts.com/v3";
        private readonly string _apiKey;

        public TracePartsService(string apiKey)
        {
            _apiKey = apiKey;
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("x-api-key", _apiKey);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        /// <summary>
        /// 1. Recherche des composants dans le catalogue TraceParts.
        /// </summary>
        /// <param name="query">Mots clés (ex: "Vanne papillon DN50")</param>
        public async Task<string> SearchProductsAsync(string query)
        {
            try
            {
                string url = $"{ApiBaseUrl}/products?q={Uri.EscapeDataString(query)}";
                HttpResponseMessage response = await _httpClient.GetAsync(url);
                response.EnsureSuccessStatusCode();
                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                return $"Erreur de recherche : {ex.Message}";
            }
        }

        /// <summary>
        /// 2. Demande la génération du fichier CAO pour un produit spécifique.
        /// </summary>
        /// <param name="productId">L'ID du produit TraceParts</param>
        /// <param name="cadFormatId">L'ID du format CAO (ex: "2" pour AutoCAD DWG 3D, à vérifier dans la doc TraceParts)</param>
        public async Task<string> RequestCadDownloadAsync(string productId, string cadFormatId = "2")
        {
            try
            {
                string url = $"{ApiBaseUrl}/products/{productId}/cad-data";
                var content = new StringContent($"{{\"cadFormatId\": \"{cadFormatId}\"}}", System.Text.Encoding.UTF8, "application/json");
                
                HttpResponseMessage response = await _httpClient.PostAsync(url, content);
                response.EnsureSuccessStatusCode();
                
                // L'API retourne généralement une URL de téléchargement ou un Job ID à vérifier
                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                return $"Erreur de téléchargement : {ex.Message}";
            }
        }
    }
}
