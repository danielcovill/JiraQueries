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

        public async Task<JiraSearchResponse> GetSearchResults(JqlSearchRequest request)
        {
            JsonContent requestJsonContent;
            HttpResponseMessage response;
            JiraSearchResponse returnData;
            var issueCollection = new List<Issue>();

            int iteration = 0;
            // Loop exists in case the maxResults we have set is smaller than the total results needed to return all issues
            do
            {
                request.startAt = request.maxResults * iteration++;
                requestJsonContent = JsonContent.Create(request,
                    typeof(JqlSearchRequest),
                    new MediaTypeHeaderValue("application/json"));
                response = await client.PostAsync("search", requestJsonContent);
                response.EnsureSuccessStatusCode();
                returnData = await response.Content.ReadFromJsonAsync<JiraSearchResponse>(
                    new JsonSerializerOptions { 
                        Converters = { new JiraDateTimeConverter() } 
                    }
                );
                issueCollection.AddRange(returnData.issues);
            } while (issueCollection.Count < returnData.total && returnData.total != 0);

            returnData.issues = issueCollection;
            returnData.total = issueCollection.Count;
            returnData.startAt = request.startAt;
            returnData.maxResults = issueCollection.Count;
            return returnData;
        }
    }
}