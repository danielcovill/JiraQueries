using System;

namespace work_charts.Models.Jira
{
    public class PointsByCategoryByTimespanReport
    {
        [CsvHelper.Configuration.Attributes.IgnoreAttribute]
        public DateTime StartDate { get; set; }
        [CsvHelper.Configuration.Attributes.IgnoreAttribute]
        public DateTime EndDate { get; set; }
        [CsvHelper.Configuration.Attributes.Index(1)]
        public int MaintenancePoints { get; set; } 
        [CsvHelper.Configuration.Attributes.Index(2)]
        public int BugPoints { get; set; } 
        [CsvHelper.Configuration.Attributes.Index(3)]
        public int TaskPoints { get; set; }
        [CsvHelper.Configuration.Attributes.Index(4)]
        public int StoryPoints { get; set; }
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