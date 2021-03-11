using System;
using System.Collections.Generic;
using CsvHelper.Configuration;

public sealed class EventFrequencyMap : ClassMap<KeyValuePair<DateTime, int>>
{
    public EventFrequencyMap()
    {
        Map(m => m.Key).Name("Date", "Key");
        Map(m => m.Value).Name("EventCount", "Value");
    }
}