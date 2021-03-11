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
                return Records.Select(r => new object[] { r.Key, r.Value }).ToArray();
            }
            set
            {
                if(value == null)
                {
                    return;
                }
                Records = value.Select(a => new KeyValuePair<DateTime, int>(
                    DateTimeOffset.FromUnixTimeSeconds(((JsonElement)a[0]).GetInt64()).UtcDateTime,
                    ((JsonElement)a[1]).EnumerateArray().ToList()[0].GetProperty("count").GetInt32()
                )).ToList();
            }
        }

        public List<KeyValuePair<DateTime, int>> Records { get; set; }

    }
}