using System.Threading.Tasks;

namespace work_charts
{
    class Program
    {
        static async Task Main(string[] args)
        {
            await JiraConnector.RunDemo();
        }
    }
}
