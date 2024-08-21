using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;

namespace SampleDocumentCreator
{
    public class ArticleExtract
    {
        public string Title { get; set; } = string.Empty;
        public string Extract { get; set; } = string.Empty;

        public static ArticleExtract DownloadWikiArticle()
        {
            var url = "https://en.wikipedia.org/w/api.php?format=json&action=query&generator=random&grnnamespace=0&prop=extracts&explaintext=1";
            using (var client = new HttpClient())
            {
                var response = client.GetAsync(url).Result;
                if (!response.IsSuccessStatusCode) throw new Exception(response.StatusCode.ToString());

                var content = response.Content.ReadAsStringAsync().Result;
                var o = JObject.Parse(content);
                var title = GetPropertyValue(o, "title");
                var extract = GetPropertyValue(o, "extract");
                if (extract.IndexOf("==") > 0) extract = extract.Substring(0, extract.IndexOf("=="));
                extract = extract.Trim();

                return new ArticleExtract()
                {
                    Title = title,
                    Extract = extract
                };
            }
        }

        private static string GetPropertyValue(JObject o, string name)
        {
            var titleIoken = o.SelectToken($"$.query.pages..{name}");
            return titleIoken.ToString();
        }
    }
}