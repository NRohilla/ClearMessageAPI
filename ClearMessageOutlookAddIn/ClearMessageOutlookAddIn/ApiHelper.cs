using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.IO;
using Newtonsoft.Json;

namespace ClearMessageOutlookAddIn
{
    public class ApiHelper
    {
        public HttpClient InitializeClient()
        {
            //string bearerKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1IjoiNjEifQ.V9wAFLJuavavvbrfCG1jWCBWwLeXYkSufx-AnhzZEPQ";

            SettingsModel settings = new SettingsModel();

            using (StreamReader sr = new StreamReader(settings.FilePath + "\\settings.json"))
            {
                var json = sr.ReadToEnd();
                settings = JsonConvert.DeserializeObject<SettingsModel>(json);
            }

            var client = new HttpClient();
            client.BaseAddress = new Uri(settings.ApiBaseUrl);
            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", settings.BearerKey);
            return client;
        }
    }
}
