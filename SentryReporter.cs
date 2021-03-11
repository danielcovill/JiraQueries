using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using CsvHelper;

namespace work_charts
{
    class SentryReporter
    {
        public SentryReporter()
        {

        }

        public void GenerateEventFrequencyReport(List<KeyValuePair<DateTime, int>> frequencyResponse, string csvOutputPath)
        {
            using (var writer = new StreamWriter(csvOutputPath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.Context.RegisterClassMap<EventFrequencyMap>();
                csv.WriteRecords(frequencyResponse);
            }
        }
    }
}