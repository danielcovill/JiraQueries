using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using work_charts.Models.Jira;

namespace work_charts
{

    class JiraReporter
    {

        public JiraReporter()
        {

        }
        /// <summary>
        /// Gives a per work item, per timespan breakdown of points completed
        /// </summary>
        public void GenerateWorkSummaryReport(JiraSearchResponse searchResponse, int daysOfHistory, int daysPerTimespan, string csvOutputPath)
        {
            var results = new List<PointsByCategoryByTimespanReport>();
            var spanEnd = DateTime.Today.AddDays(1).AddMilliseconds(-1);
            while (daysOfHistory > 0)
            {
                results.Add(new PointsByCategoryByTimespanReport()
                {
                    StartDate = spanEnd.AddDays(-daysPerTimespan),
                    EndDate = spanEnd,
                    BugPoints = Convert.ToInt32(searchResponse.GetTickets(
                        ticketType: "Bug",
                        resolvedStartDate: spanEnd.AddDays(-7),
                        resolvedEndDate: spanEnd)
                            .Select(issue => issue.fields)
                            .Sum(field => field.storyPoints).Value),
                    MaintenancePoints = Convert.ToInt32(searchResponse.GetTickets(
                        ticketType: "Maintenance",
                        resolvedStartDate: spanEnd.AddDays(-7),
                        resolvedEndDate: spanEnd)
                            .Select(issue => issue.fields)
                            .Sum(field => field.storyPoints).Value),
                    StoryPoints = Convert.ToInt32(searchResponse.GetTickets(
                        ticketType: "Story",
                        resolvedStartDate: spanEnd.AddDays(-7),
                        resolvedEndDate: spanEnd)
                            .Select(issue => issue.fields)
                            .Sum(field => field.storyPoints).Value),
                    TaskPoints = Convert.ToInt32(searchResponse.GetTickets(
                        ticketType: "Task",
                        resolvedStartDate: spanEnd.AddDays(-7),
                        resolvedEndDate: spanEnd)
                            .Select(issue => issue.fields)
                            .Sum(field => field.storyPoints).Value),
                });
                spanEnd = spanEnd.AddDays(-daysPerTimespan);
                daysOfHistory -= daysPerTimespan;
            }
            using (var writer = new StreamWriter(csvOutputPath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(results.OrderBy(x => x.EndDate));
            }
        }

        /// <summary>
        /// Gives a per-engineer summary of points completed over the last 10 weeks by providing 10 and 2 week averages as
        /// well as respective point totals for the previous week
        /// </summary>
        public void GenerateTeamOutputReport(JiraSearchResponse searchResponse, List<User> engineers, string csvOutputPath)
        {
            var results = new List<EngineerOutputReport>();
            foreach (var engineer in engineers.OrderBy(a => a.name))
            {
                results.Add(new EngineerOutputReport()
                {
                    Name = engineer.displayName,
                    TenWeekAveragePoints = searchResponse.GetTickets(
                        accountId: engineer.accountId,
                        resolvedStartDate: DateTime.Now.AddDays(-70),
                        resolvedEndDate: DateTime.Now)
                            .Select(issue => issue.fields)
                            .Sum(field => field.storyPoints).Value / 10,
                    TwoWeekAveragePoints = searchResponse.GetTickets(
                        accountId: engineer.accountId,
                        resolvedStartDate: DateTime.Now.AddDays(-14),
                        resolvedEndDate: DateTime.Now)
                            .Select(issue => issue.fields)
                            .Sum(field => field.storyPoints).Value / 2,
                    PriorWeekPoints = searchResponse.GetTickets(
                        accountId: engineer.accountId,
                        resolvedStartDate: DateTime.Now.AddDays(-7),
                        resolvedEndDate: DateTime.Now)
                            .Select(issue => issue.fields)
                            .Sum(field => field.storyPoints).Value
                });
            }
            using (var writer = new StreamWriter(csvOutputPath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(results);
            }
        }

        public void GenerateBugReport(JiraSearchResponse bugSearchResponse, JiraSearchResponse allSearchResponse, int daysOfHistory, int daysPerTimespan, string csvOutputPath)
        {
            var results = new List<BugsReport>();
            var spanEnd = DateTime.Today.AddDays(1).AddMilliseconds(-1);
            while (daysOfHistory > 0)
            {
                var createdBugsCollection = bugSearchResponse.GetTickets(
                    createdStartDate: spanEnd.AddDays(-daysPerTimespan),
                    createdEndDate: spanEnd);
                var createdBugsCount = createdBugsCollection.Count();
                var completedStoryMaintCount = allSearchResponse.GetTickets(
                    resolvedStartDate: spanEnd.AddDays(-daysPerTimespan),
                    resolvedEndDate: spanEnd).Where(ticket => ticket.fields.issuetype.name == "Bug" || ticket.fields.issuetype.name == "Story").Count();
                var completedStoryMaintPoints = allSearchResponse.GetTickets(
                    resolvedStartDate: spanEnd.AddDays(-daysPerTimespan),
                    resolvedEndDate: spanEnd).Where(ticket => ticket.fields.storyPoints != null && (ticket.fields.issuetype.name == "Bug" || ticket.fields.issuetype.name == "Story")).Sum(ticket => ticket.fields.storyPoints).Value;
                var escapedBugCount = createdBugsCollection.Where(bug =>
                    bug.fields.exposure != null && bug.fields.exposure.value == "Production").Count();

                results.Add(new BugsReport()
                {
                    StartDate = spanEnd.AddDays(-daysPerTimespan),
                    EndDate = spanEnd,
                    BugsEscaped = escapedBugCount,
                    BugsCaught = createdBugsCount - escapedBugCount,
                    BugsPerTenStoryMaintTickets = ((double)createdBugsCount / completedStoryMaintCount) * 10,
                    BugsPerStoryMaintPoint = (double)createdBugsCount / completedStoryMaintPoints,
                    Regressions = createdBugsCollection.Where(bug => bug.fields.regressions != null && bug.fields.regressions.Any(regression => regression.value.Equals("Is Regression"))).Count(),
                });
                spanEnd = spanEnd.AddDays(-daysPerTimespan);
                daysOfHistory -= daysPerTimespan;
            }
            using (var writer = new StreamWriter(csvOutputPath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(results.OrderBy(x => x.EndDate));
            }
        }

    }
}