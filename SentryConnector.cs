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

        public async Task<List<KeyValuePair<DateTime, int>>> GetEventFrequency(SentryRequest request) 
        {
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
            var result = System.Text.Json.JsonSerializer.Deserialize<SentryEventFrequencyResponse>(await response.Content.ReadAsByteArrayAsync());
            return result.Records;
        }

        public async Task<SentryFrequentEvent[]> GetFrequentEvents(SentryRequest request)
        {
            var requestString = new StringBuilder("eventsv2/");
            requestString.Append($"?query={Uri.EscapeUriString(request.query)}");
            requestString.Append($"&statsPeriod={Uri.EscapeUriString(request.statsPeriod)}");
            requestString.Append($"&sort={Uri.EscapeUriString(request.sort)}");
            foreach(string field in request.fields) 
            {
                requestString.Append($"&field={Uri.EscapeUriString(field)}");
            }
            var response = await client.GetAsync(requestString.ToString());
            response.EnsureSuccessStatusCode();
            var result = System.Text.Json.JsonSerializer.Deserialize<SentryFrequentEventsResponse>(await response.Content.ReadAsByteArrayAsync());
            return result.data;
        }
    }
}