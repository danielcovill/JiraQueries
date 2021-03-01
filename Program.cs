using System;
using System.IO;
using System.Threading.Tasks;

namespace work_charts
{
    class Program
    {
        public static async Task Main(string[] args)
        {
            var reports = new string[]
            {
                "Team Output - 10wk, 2wk, and 1wk output per Engineer",
                "Weekly Points Summary - Broken down by work item type",
                "Bugs Summary - Overview of regressions, escapes, etc.",
                "Stalled Summary - Overview of stalled and blocked tickets"
            };

            var requestedReport = selectReport(reports);

            var reporter = new JiraReporter();
            var outputPath = Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop),
                $"jiraReport-{DateTime.Now.ToString("yyyy-MM-ddTHHmm-ss")}.csv");

            switch (requestedReport)
            {
                case "Team Output - 10wk, 2wk, and 1wk output per Engineer":
                {
                    var jqlRestRequest = new JqlSearchRequest(File.ReadAllText(Path.Combine(@"Queries", "Done previous 10wk.jql")));
                    var searchResult = await JiraConnector.Instance.GetSearchResults(jqlRestRequest);
                    var engineeringGroup = await JiraConnector.Instance.GetGroup(new JiraGroupRequest("engineers"));
                    var engineers = engineeringGroup.users;
                    reporter.GenerateTeamOutputReport(searchResult, engineers,
                        Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                    break;
                }
                case "Weekly Points Summary - Broken down by work item type":
                {
                    var jqlRestRequest = new JqlSearchRequest(File.ReadAllText(Path.Combine(@"Queries", "Done previous 10wk.jql")));
                    var searchResult = await JiraConnector.Instance.GetSearchResults(jqlRestRequest);
                    reporter.GenerateWorkSummaryReport(searchResult, 70, 7,
                        Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                    break;
                }
                case "Bugs Summary - Overview of regressions, escapes, etc.":
                {
                    var jqlBugRestRequest = new JqlSearchRequest(File.ReadAllText(Path.Combine(@"Queries", "Bugs created 10wk.jql")));
                    var jqlRestRequest = new JqlSearchRequest(File.ReadAllText(Path.Combine(@"Queries", "Done previous 10wk.jql")));
                    var jiraQueries = new Task<JiraSearchResponse>[2] {
                        JiraConnector.Instance.GetSearchResults(jqlBugRestRequest),
                        JiraConnector.Instance.GetSearchResults(jqlRestRequest),
                    };
                    Task.WaitAll(jiraQueries);
                    reporter.GenerateBugReport(jiraQueries[0].Result, jiraQueries[1].Result, 70, 7,
                        Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                    break;
                }
                case "Stalled Summary - Overview of stalled and blocked tickets":
                {
                    var jqlRestRequest = new JqlSearchRequest(File.ReadAllText(Path.Combine(@"Queries", "Blocked and Stalled.jql")));
                    var searchResult = await JiraConnector.Instance.GetSearchResults(jqlRestRequest);
                    reporter.GenerateStalledReport(searchResult,
                        Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                    break;
                }
            }
            Console.WriteLine($"Report generated: {outputPath}");
        }
        private static string selectReport(string[] reports)
        {
            if (reports.Length == 0)
            {
                throw new Exception("No reports available");
            }

            Console.WriteLine("Select a report ('q' to exit):");
            var iterator = 0;
            foreach (var report in reports)
            {
                Console.Write(++iterator + ". ");
                Console.WriteLine(report);
            }
            var selection = 0;
            while (true)
            {
                var keyPress = Console.ReadLine();
                if (keyPress == "q")
                {
                    Environment.Exit(0);
                }
                if (Int32.TryParse(keyPress, out selection) && selection > 0 && selection <= reports.Length)
                {
                    return reports[selection - 1];
                }
            }
        }
    }
}