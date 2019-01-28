using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;

namespace ClearMessageOutlookAddIn
{
    public class ApiHelper
    {
        //private const string apiBaseUrl = "https://private-f8e32-clearapiprivate.apiary-proxy.com/";
        private const string apiBaseUrl = "https://api.clearmessage.com/";
        private string bearerKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1IjoiNjEifQ.V9wAFLJuavavvbrfCG1jWCBWwLeXYkSufx-AnhzZEPQ";

        public HttpClient InitializeClient()
        {
            var client = new HttpClient();   
            client.BaseAddress = new Uri(apiBaseUrl);
            client.DefaultRequestHeaders.Clear(); 
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization =new AuthenticationHeaderValue("Bearer", bearerKey);
            return client;
        }
    }
}
