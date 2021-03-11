namespace work_charts.Models.Sentry
{
    public class SentryRequest
    {
        public string query { get; set; }
        public string statsPeriod { get; set; }
        public string sort { get; set; }

        public int perPage { get; set; }
        public string[] fields { get; set; }
        public string interval { get; set; }
    }
}