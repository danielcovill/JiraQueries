using System;

namespace work_charts.Models
{
    public class StalledTicket
    {
        public string Id { get; set; }
        public string Status { get; set; }
        public string Summary { get; set; }
        public string Type { get; set; }
        public string Link { get; set; }
        public int DaysStalled { get; set; }
        public DateTime Created { get; set; }
        public DateTime Updated { get; set; }
        public string Creator { get; set; }
        public string Assignee { get; set; }
        public string UpdatedBy { get; set; }
        public int StoryPoints { get; set; }
        public bool IsRegression { get; set; }
        public string Epic { get; set; }
    }
}