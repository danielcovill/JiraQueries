using System;

namespace work_charts.Models.Jira
{
    public class BugsReport
    {
        [CsvHelper.Configuration.Attributes.IgnoreAttribute]
        public DateTime StartDate { get; set; }
        [CsvHelper.Configuration.Attributes.IgnoreAttribute]
        public DateTime EndDate { get; set; }
        [CsvHelper.Configuration.Attributes.Index(1)]
        public int BugsEscaped { get; set; } 
        [CsvHelper.Configuration.Attributes.Index(2)]
        public int BugsCaught { get; set; } 
        [CsvHelper.Configuration.Attributes.Index(3)]
        public double BugsPerTicket { get; set; }
        [CsvHelper.Configuration.Attributes.Index(4)]
        public double BugsPerPoint {get; set; }
        [CsvHelper.Configuration.Attributes.Index(5)]
        public int Regressions { get; internal set; }
        [CsvHelper.Configuration.Attributes.Index(0)]
        public string DateRange
        {
            get
            {
                return $"{this.StartDate.ToString("dd-MMM")} - {this.EndDate.ToString("dd-MMM")}";
            }
        }

    }
}