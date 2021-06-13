namespace DocXGenerator
{
    using System;
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.Http;
    using Microsoft.AspNetCore.Http;
    using Microsoft.Extensions.Logging;
    using Newtonsoft.Json;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.Net.Http;

    public static class Function1
    {
        [FunctionName("GenerateDocument")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
        {
            string templateUrl = null;
            Dictionary<string, string> dictionary = null;

            using (StreamReader streamReader = new StreamReader(req.Body))
            {
                var json = await streamReader.ReadToEndAsync();

                if (String.IsNullOrWhiteSpace(json))
                {
                    return new OkObjectResult("No POST data provided.");
                }

                dictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
            }

            if (dictionary.ContainsKey("_TEMPLATE_URL"))
            {
                templateUrl = dictionary["_TEMPLATE_URL"];
                dictionary.Remove("_TEMPLATE_URL");
            }

            if (dictionary == null || dictionary.Count == 0)
            {
                return new OkObjectResult("Could not convert JSON into dictionary.");
            }

            var stream = await DownloadTemplate(templateUrl);

            using (MemoryStream stream2 = new MemoryStream())
            {
                stream.CopyTo(stream2);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream2, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;

                    foreach (var textChild in body.Descendants<Text>())
                    {
                        if (textChild.InnerText.Contains("((") && textChild.InnerText.Contains("))"))
                        {
                            TryInterpolate(dictionary, textChild);
                        }
                    }
                }

                return new FileContentResult(stream2.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            }
        }

        static async Task<Stream> DownloadTemplate(string url)
        {
            using (HttpClient httpClient = new HttpClient())
            {
                var response = await httpClient.GetAsync(url);
                return await response.Content.ReadAsStreamAsync();
            }
        }

        static void TryInterpolate(Dictionary<string, string> dictionary, Text text)
        {
            if (String.IsNullOrWhiteSpace(text.Text))
            {
                return;
            }

            foreach (var entry in dictionary)
            {
                if (text.Text.Equals($"(({entry.Key}))", StringComparison.OrdinalIgnoreCase))
                {
                    text.Text = text.Text.Replace($"(({entry.Key}))", entry.Value);
                    return;
                }
            }
        }
    }
}
