using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using CsvHelper.Configuration;

namespace work_charts.Models.Sentry
{
    public class SentryFrequentEventsResponse
    {
        public SentryFrequentEvent[] data { get; set; }
    }
    public class SentryFrequentEvent
    {
        public string title { get; set; }
        public int count { get; set; }
    }
}