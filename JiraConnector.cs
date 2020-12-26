using System;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Extensions.Configuration;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Net.Http.Json;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace work_charts
{
    public sealed class JiraConnector
    {
        private Uri baseAddress = new Uri("https://smartwyre.atlassian.net/rest/api/latest/");
        private readonly HttpClient client;
        private static readonly Lazy<JiraConnector> lazyJiraConnector = new Lazy<JiraConnector>(() => new JiraConnector());
        public static JiraConnector Instance
        {
            get
            {
                return lazyJiraConnector.Value;
            }
        }
        public string results { get; set; }

        public JiraConnector()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("connections.json");
            var configuration = builder.Build();

            this.client = new HttpClient();
            client.BaseAddress = baseAddress;

            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue(AuthenticationSchemes.Basic.ToString(),
                Convert.ToBase64String(System.Text.Encoding.UTF8
                    .GetBytes($"{configuration["jiraUser"]}:{configuration["jiraApiKey"]}")));
        }

        public async Task<string> RunSearch(JqlSearchRequest request)
        {
            var requestString = JsonSerializer.Serialize(request, typeof(JqlSearchRequest));
            var jsonContent = JsonContent.Create(request,
                typeof(JqlSearchRequest),
                new MediaTypeHeaderValue("application/json"));
            var response = await client.PostAsync("search", jsonContent);
            response.EnsureSuccessStatusCode();
            var deserializeOptions = new JsonSerializerOptions
            {
                Converters =
                {
                    new JiraDateTimeConverter()
                }
            };
            JiraSearchResponse result = await response.Content.ReadFromJsonAsync<JiraSearchResponse>(deserializeOptions);
            List<Issue> resultSet = result.issues;
            while(result.total > result.maxResults) 
            {
                //repeat search for the next set of paged results and append total
            }

            //TODO: Build out paging mechanism
            //if total is greater than maxresults, update startAt by whatever multiple is appropriate and re-request until done
            //since I've already parsed the thing at this point, I should probably hand back a parsed response instead of a string
            return "";
        }
    }
}