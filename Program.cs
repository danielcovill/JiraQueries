using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace work_charts
{
    class Program
    {
        public static async Task Main(string[] args)
        {
            var reports = new Dictionary<string, string>()
            {
                { "Team Output - 10wk, 2wk, and 1wk output per Engineer", "Done previous 10wk.jql" }
            };

            var requestedReport = selectReport(reports);

            var reporter = new JiraReporter();
            var outputPath = Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop),
                $"jiraReport-{DateTime.Now.ToString("yyyy-MM-ddTHHmm-ss")}");
            var jqlRestRequest = new JqlSearchRequest(File.ReadAllText(Path.Combine(@".\Queries", requestedReport.Value)));
            var searchResult = await JiraConnector.Instance.GetSearchResults(jqlRestRequest);
            var engineeringGroup = await JiraConnector.Instance.GetGroup(new JiraGroupRequest("engineers"));
            var engineers = engineeringGroup.users;

            switch (requestedReport.Key)
            {
                case "Team Output - 10wk, 2wk, and 1wk output per Engineer":
                    outputPath = Path.ChangeExtension(outputPath, "csv");
                    reporter.GenerateTeamOutputReport(searchResult, engineers,
                        Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                    break;
            }
            Console.WriteLine($"Report generated: {outputPath}");
        }
        private static KeyValuePair<string,string> selectReport(Dictionary<string, string> reports)
        {
            if (reports.Count == 0)
            {
                throw new Exception("No reports available");
            }

            Console.WriteLine("Select a report ('q' to exit):");
            var iterator = 0;
            foreach (var report in reports)
            {
                Console.Write(++iterator + ". ");
                Console.WriteLine(report.Key);
            }
            var selection = 0;
            while (true)
            {
                var keyPress = Console.ReadLine();
                if (keyPress == "q")
                {
                    Environment.Exit(0);
                }
                if (Int32.TryParse(keyPress, out selection) && selection > 0 && selection <= reports.Count)
                {
                    return reports.ElementAt(selection - 1);
                }
            }
        }
    }
}