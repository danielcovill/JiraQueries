using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Text.Json;

namespace work_charts.Models.Sentry
{
    public class SentryEventFrequencyResponse
    {
        public object[][] data
        {
            get 
            {
                if(Records == null)
                {
                    return null;
                }
                return Records.Select(r => new object[] { r.Date, r.EventCount}).ToArray();
            }
            set
            {
                if(value == null)
                {
                    return;
                }
                Records = value.Select(a => new SentryDateEventCount
                {
                    Date = DateTimeOffset.FromUnixTimeSeconds(((JsonElement)a[0]).GetInt64()).UtcDateTime,
                    EventCount = ((JsonElement)a[1]).EnumerateArray().ToList()[0].GetProperty("count").GetInt32()
                }).ToList();
            }
        }

        public List<SentryDateEventCount> Records { get; set; }

    }
    public class SentryDateEventCount
    {
        public DateTime Date { get; set; }
        public int EventCount { get; set; }
    }
    public class SentryCount {
        public int count { get; set; }
    }
}