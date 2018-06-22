using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace TranslationEndPoint
{

    // Models will be kept here for now.
    class detectedLanguage
    {
        public String language { get; set; }
        public float score { get; set; }
    }
    class translations
    {
        public String text { get; set; }
        public String to { get; set; }
    }
    class TranslationObject
    {
        public detectedLanguage detectedLanguage { get; set; }
        public List<translations> translations { get; set; }
    }
    class TranslationAPI
    {
        static string host = "https://api.cognitive.microsofttranslator.com";
        static string path = "/translate?api-version=3.0";
        string params_from = "&from=";
        string params_to = "&to=";

        static string uri = host + path;

        static string key = ""; // key

        public TranslationAPI(String from, String to) {
            params_from += from;
            params_to += to;
        }
        public static Task<String> translate(String from, String to, String paragraph) {
            TranslationAPI trans = new TranslationAPI(from, to);
            return trans.asyncTranslationQuery(paragraph);
        }

        async Task<String> asyncTranslationQuery(String paragraph)
        {
            
            System.Object[] body = new System.Object[] { new { Text = paragraph } };
            var requestBody = JsonConvert.SerializeObject(body);

            
            using (var client = new HttpClient())
            using (var request = new HttpRequestMessage())
            {
                request.Method = HttpMethod.Post;
                request.RequestUri = new Uri(uri+(params_from + params_to));
                request.Content = new StringContent(requestBody, Encoding.UTF8, "application/json");
                request.Headers.Add("Ocp-Apim-Subscription-Key", key);

                var response = await client.SendAsync(request).ConfigureAwait(false);
                var responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                try
                {
                    List<TranslationObject> resultObject = JsonConvert.DeserializeObject<List<TranslationObject>>(responseBody);
                    String result = resultObject[0].translations[0].text;
                    return result;
                }
                catch (Exception e) {
                    Console.WriteLine("error " + e.Message);
                    Console.WriteLine("error p " + paragraph);
                    Console.WriteLine("error p " + JsonConvert.DeserializeObject(responseBody));
                    Console.ReadLine();
                    return "";
                }

            }
        }

    }
}
