namespace work_charts.Models
{
    public class EngineerOutputReport
    {
        public string Name { get; set; }
        public double TenWeekAveragePoints { get; set; }

        public double TwoWeekAveragePoints { get; set; }
        public double PriorWeekPoints{ get; set; }
    }
}