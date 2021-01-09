using System;
namespace work_charts
{
    class ReportRequest
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Query { get; set; }
    }
}