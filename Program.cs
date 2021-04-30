using System;
using System.IO;
using System.Threading.Tasks;
using work_charts.Models.Jira;
using work_charts.Models.Sentry;

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
                "Bugs by Team - Bugs generated per Engineer",
                "Sentry Events Frequency",
                "Sentry most frequent events -7d"
            };
            while (true)
            {
                var requestedReport = selectReport(reports);

                var outputPath = Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop),
                    $"Report-{DateTime.Now.ToString("yyyy-MM-ddTHHmm-ss")}.csv");

                switch (requestedReport)
                {
                    case "Team Output - 10wk, 2wk, and 1wk output per Engineer":
                    {
                        var reporter = new JiraReporter();
                        var jqlRestRequest = new JqlSearchRequest(File.ReadAllText(Path.Combine(@"Models", @"Jira", @"Queries", "Done previous 10wk.jql")));
                        var searchResult = await JiraConnector.Instance.GetSearchResults(jqlRestRequest);
                        var engineeringGroup = await JiraConnector.Instance.GetGroup(new JiraGroupRequest("engineers"));
                        var engineers = engineeringGroup.users;
                        reporter.GenerateTeamOutputReport(searchResult, engineers,
                            Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                        break;
                    }
                    case "Weekly Points Summary - Broken down by work item type":
                    {
                        var reporter = new JiraReporter();
                        var jqlRestRequest = new JqlSearchRequest(File.ReadAllText(Path.Combine(@"Models", @"Jira", @"Queries", "Done previous 10wk.jql")));
                        var searchResult = await JiraConnector.Instance.GetSearchResults(jqlRestRequest);
                        reporter.GenerateWorkSummaryReport(searchResult, 70, 7,
                            Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                        break;
                    }
                    case "Bugs Summary - Overview of regressions, escapes, etc.":
                    {
                        var reporter = new JiraReporter();
                        var jqlBugRestRequest = new JqlSearchRequest(File.ReadAllText(Path.Combine(@"Models", @"Jira", @"Queries", "Bugs created 10wk.jql")));
                        var jqlRestRequest = new JqlSearchRequest(File.ReadAllText(Path.Combine(@"Models", @"Jira", @"Queries", "Done previous 10wk.jql")));
                        var jiraQueries = new Task<JiraSearchResponse>[2] {
                            JiraConnector.Instance.GetSearchResults(jqlBugRestRequest),
                            JiraConnector.Instance.GetSearchResults(jqlRestRequest),
                        };
                        Task.WaitAll(jiraQueries);
                        reporter.GenerateBugReport(jiraQueries[0].Result, jiraQueries[1].Result, 70, 7,
                            Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                        break;
                    }
                    case "Bugs by Team - Bugs generated per Engineer":
                    {
                        var reporter = new JiraReporter();
                        var jqlBugRestRequest = new JqlSearchRequest(File.ReadAllText(Path.Combine(@"Models", @"Jira", @"Queries", "Bugs completed 10wk.jql")));
                        var searchResult = await JiraConnector.Instance.GetSearchResults(jqlBugRestRequest);
                        break;
                    }
                    case "Sentry Events Frequency":
                    {
                        var reporter = new SentryReporter();
                        var sentryRequest = new SentryRequest()
                        {
                            interval = "1d",
                            sort = "-timestamp",
                            statsPeriod = "70d",
                            query = "event.type:error",
                            fields = new[] { "count()", "timestamp" }
                        };
                        var eventFrequencyResult = await SentryConnector.Instance.GetEventFrequency(sentryRequest);
                        reporter.GenerateEventFrequencyReport(eventFrequencyResult,
                            Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                        break;
                    }
                    case "Sentry most frequent events -7d":
                    {
                        var reporter = new SentryReporter();
                        var sentryRequest = new SentryRequest()
                        {
                            sort = "-count",
                            statsPeriod = "7d",
                            query = "event.type:error !is:ignored",
                            fields = new[] { "count()", "title" }
                        };
                        var frequentEventsResult = await SentryConnector.Instance.GetFrequentEvents(sentryRequest);
                        reporter.GenerateFrequentEventsReport(frequentEventsResult,
                            Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                        break;
                    }
                    default:
                    {
                        throw new Exception("Requested report does not exist");
                    }
                }
                Console.WriteLine($"Report generated: {outputPath}");
            }
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