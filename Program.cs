using System;
using System.IO;
using System.Threading.Tasks;

namespace work_charts
{
    class Program
    {
        public static async Task Main(string[] args)
        {
            var request = new JqlSearchRequest(File.ReadAllText("./Queries/completed_last_week.jql"));
            var result = await JiraConnector.Instance.RunSearch(request);

            //Temporary
            File.WriteAllText(Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "out.json"), result);
        }
    }
}