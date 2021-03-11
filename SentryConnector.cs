using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Extensions.Configuration;
using work_charts.Models.Sentry;

namespace work_charts
{
    public sealed class SentryConnector
    {
        private Uri baseAddress = new Uri("https://sentry.io/api/0/organizations/smartwyre/");
        private readonly HttpClient client;
        private static readonly Lazy<SentryConnector> lazySentryConnector = new Lazy<SentryConnector>(() => new SentryConnector());

        public static SentryConnector Instance
        {
            get
            {
                return lazySentryConnector.Value;
            }
        }

        public SentryConnector()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("connections.json");
            var configuration = builder.Build();

            this.client = new HttpClient();
            client.BaseAddress = baseAddress;
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", configuration["sentryApiKey"]);
        }

        public async Task<Dictionary<DateTime, int>> GetEventFrequency(SentryRequest request) 
        {
            var result = new Dictionary<DateTime, int>();
            var requestString = new StringBuilder("events-stats/");
            requestString.Append($"?query={Uri.EscapeUriString(request.query)}");
            requestString.Append($"&interval={Uri.EscapeUriString(request.interval)}");
            requestString.Append($"&statsPeriod={Uri.EscapeUriString(request.statsPeriod)}");
            requestString.Append($"&sort={Uri.EscapeUriString(request.sort)}");
            foreach(string field in request.fields) 
            {
                requestString.Append($"&field={Uri.EscapeUriString(field)}");
            }
            
            var response = await client.GetAsync(requestString.ToString());
            response.EnsureSuccessStatusCode();
            string x = await response.Content.ReadAsStringAsync();
            var stupidResult = System.Text.Json.JsonSerializer.Deserialize<SentryEventFrequencyResponse>(x);
            return result;
        }
    }
}