using System;
using System.IO;
using System.Threading.Tasks;

namespace work_charts
{
    class Program
    {
        static async Task Main(string[] args)
        {
            //await JiraConnector.RunDemo();
            Console.WriteLine(File.ReadAllText("./Queries/completed_last_week.jql"));
            Console.WriteLine();
            //await JiraConnector.RunQuery(query);
        }
    }
}