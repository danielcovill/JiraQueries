using System;
using System.IO;
using System.Threading.Tasks;

namespace work_charts
{
    class Program
    {
        public static async Task Main(string[] args)
        {
            string requestedQuery = null;
            string requestedReport = null;
            var jqlQueries = Directory.GetFiles(@".\Queries", "*.jql");
            string[] reports = {
                "Summary",
                "WeeklyBreakdown",
                "TeamOutput"
            };

            //pick report and query in order
            while (requestedReport == null)
            {
                if (requestedQuery == null)
                {
                    requestedQuery = selectQuery(jqlQueries);
                }
                if (requestedReport == null)
                {
                    requestedReport = selectReport(reports);
                    if (requestedReport == null)
                    {
                        requestedQuery = null;
                    }
                }
            }
            var reporter = new JiraReporter();
            var outputPath = Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop),
                $"{Path.GetFileNameWithoutExtension(requestedQuery)}-{requestedReport}-{DateTime.Now.ToString("yyyy-MM-ddTHHmm-ss")}");
            var jqlRestRequest = new JqlSearchRequest(File.ReadAllText(requestedQuery));
            var searchResult = await JiraConnector.Instance.GetSearchResults(jqlRestRequest);
            var engineeringGroup = await JiraConnector.Instance.GetGroup(new JiraGroupRequest("engineers")); 
            var engineers = engineeringGroup.users;

            switch (requestedReport)
            {
                case "Summary":
                    reporter.GenerateTicketListSummary(searchResult, engineers,
                        Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath), jqlRestRequest.jql);
                    break;
                case "TeamOutput":
                    outputPath = Path.ChangeExtension(outputPath, "csv");
                    reporter.GenerateTeamOutputReport(searchResult, engineers, 
                        Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                    break;
                case "WeeklyBreakdown":
                default:
                    reporter.GenerateTimespanBreakdown(searchResult, engineers, DateTime.Today.AddDays(-7), DateTime.Today,
                        Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), outputPath));
                    break;
            }
            Console.WriteLine($"Report generated: {outputPath}");
        }

        private static string selectReport(string[] reports)
        {
            if (reports.Length == 0)
            {
                throw new Exception("No reports available");
            }

            Console.WriteLine("Select a report ('q' to exit, 'b' to go back):");
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
                if (keyPress == "b")
                {
                    return null;
                }
                if (Int32.TryParse(keyPress, out selection) && selection > 0 && selection <= reports.Length)
                {
                    return reports[selection - 1];
                }
            }
        }

        private static string selectQuery(string[] jqlQueries)
        {
            if (jqlQueries.Length == 0)
            {
                Console.WriteLine("No jql files in the 'Queries' directory. Exiting.");
                Environment.Exit(0);
            }
            while (true)
            {
                Console.WriteLine("Select a query ('q' to exit):");
                int iterator = 0;
                foreach (var jqlQuery in jqlQueries)
                {
                    Console.Write(++iterator + ". ");
                    Console.WriteLine(Path.GetFileNameWithoutExtension(jqlQuery));
                }
                int selection = 0;

                var keyPress = Console.ReadLine();
                if (keyPress == "q")
                {
                    Environment.Exit(0);
                }
                if (Int32.TryParse(keyPress, out selection) && selection > 0 && selection < jqlQueries.Length)
                {
                    return jqlQueries[selection - 1];
                }
            }
        }
    }
}