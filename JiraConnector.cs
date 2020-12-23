using System;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Extensions.Configuration;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Net.Http.Json;

namespace work_charts
{
    public static class JiraConnector
    {
        public static IConfiguration configuration { get; set; }
        private static readonly HttpClient client = new HttpClient()
        {
            BaseAddress = new Uri("https://smartwyre.atlassian.net/rest/api/latest/")
        };

        private static void PrepQuery() 
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("connections.json");
            configuration = builder.Build();

            client.DefaultRequestHeaders.Authorization = 
                new AuthenticationHeaderValue(AuthenticationSchemes.Basic.ToString(), 
                Convert.ToBase64String(System.Text.Encoding.UTF8
                    .GetBytes($"{configuration["jiraUser"]}:{configuration["jiraApiKey"]}")));
        }

        public static async Task RunDemo()
        {
            PrepQuery();
            var response = await client.GetAsync("issue/SMART-1911");
            response.EnsureSuccessStatusCode();
            Console.WriteLine(await response.Content.ReadAsStringAsync());
        }

        public static async Task RunSearch(string jql, string[] fields, string expansionType)
        {
            PrepQuery();
            var request = new JqlSearchRequest(jql, fields);
            var requestString = JsonSerializer.Serialize(request, typeof(JqlSearchRequest));
            var jsonContent = JsonContent.Create(request, 
                typeof(JqlSearchRequest), 
                new MediaTypeHeaderValue("application/json"));
            var response = await client.PostAsync("search", jsonContent);
        }
    }
}