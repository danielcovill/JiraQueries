using System;
using System.IO;
using System.Threading.Tasks;

namespace work_charts
{
    class Program
    {
        public static async Task Main(string[] args)
        {
            //build out the interface with options
            var completedLastWeekRequest = new JqlSearchRequest(File.ReadAllText("./Queries/completed_last_week.jql"));
            var completedLastWeekResult = await JiraConnector.Instance.GetSearchResults(completedLastWeekRequest);
            var reporter = new JiraReporter();
            var xlsxOutputPath = Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), 
                $"Previous-Week-{DateTime.Now.ToString("yyyy-MM-ddTHHmm-ssf")}.xlsx");
            Console.WriteLine(xlsxOutputPath);
            reporter.GenerateTicketListSummary(completedLastWeekResult, DateTime.Today.AddDays(-7), DateTime.Today,
                Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), xlsxOutputPath));
        }
    }
}