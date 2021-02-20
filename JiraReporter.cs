using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;

namespace work_charts
{

    class JiraReporter
    {

        public JiraReporter()
        {

        }
        /// <summary>
        /// Gets a breakdown of ticket counts, points, and ratios among ticket types for a given list of tickets
        /// </summary>
        /// <param name="issues">The list of issues over which to compute</param>
        /// <param name="xlsxOutputPath">If included output creates an xlsx document at the specified path. Otherwise output goes to the console.</param>
        public void GenerateTeamOutputReport(JiraSearchResponse searchResponse, List<User> engineers, string csvOutputPath)
        {
            var results = new List<EngineerOutputReport>();
            foreach (var engineer in engineers)
            {
                results.Add(new EngineerOutputReport()
                {
                    Name = engineer.displayName,
                    TenWeekAveragePoints = searchResponse.GetPointTotal(null, engineer.accountId, DateTime.Now.AddDays(-70), DateTime.Now) / 10,
                    TwoWeekAveragePoints = searchResponse.GetPointTotal(null, engineer.accountId, DateTime.Now.AddDays(-14), DateTime.Now) / 2,
                    PriorWeekPoints = searchResponse.GetPointTotal(null, engineer.accountId, DateTime.Now.AddDays(-7), DateTime.Now)
                });
            }
            using (var writer = new StreamWriter(csvOutputPath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(results);
            }
        }


    }
}