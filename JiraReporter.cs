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
            foreach (var engineer in engineers)
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
                    bug.fields.environment != null && bug.fields.environment.Contains("production", StringComparison.InvariantCultureIgnoreCase) ||
                    bug.fields.environment != null && bug.fields.environment.Contains("all", StringComparison.InvariantCultureIgnoreCase) ||
                    bug.fields.environment != null && bug.fields.environment.Contains("prod", StringComparison.InvariantCultureIgnoreCase)).Count();

                results.Add(new BugsReport()
                {
                    StartDate = spanEnd.AddDays(-daysPerTimespan),
                    EndDate = spanEnd,
                    BugsEscaped = escapedBugCount,
                    BugsCaught = createdBugsCount - escapedBugCount,
                    BugsPerTenStoryMaintTickets = ((double)createdBugsCount / completedStoryMaintCount) * 10,//FIXME - needs to only count completed story/maint tickets
                    BugsPerStoryMaintPoint = (double)createdBugsCount / completedStoryMaintPoints,//FIXME - needs to only count completed story/maint ticket points
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

        public void GenerateStalledReport(JiraSearchResponse stalledSearchResponse, string csvOutputPath) 
        {
            var results = stalledSearchResponse.issues.Select(ticket => new StalledTicket() {
                Assignee = ticket.fields.assignee != null ? ticket.fields.assignee.displayName : "",
                Created = ticket.fields.created,
                Creator = ticket.fields.creator.displayName,
                DaysStalled = (int)(DateTime.Now - ticket.fields.updated).TotalDays,
                Epic = ticket.fields.epicLink,
                Link = $"https://smartwyre.atlassian.net/browse/{ticket.key}",
                Id = ticket.key,
                IsRegression = (ticket.fields.regressions != null && ticket.fields.regressions.Any(regression => regression.value.Equals("Is Regression"))),
                Status = ticket.fields.status.name,
                StoryPoints = ticket.fields.storyPoints.HasValue ? (int)ticket.fields.storyPoints.Value : 0,
                Summary = ticket.fields.summary,
                Type = ticket.fields.issuetype.name,
                Updated = ticket.fields.updated,
                UpdatedBy = "not implemented",//TODO: Implement me
            });
            
            using (var writer = new StreamWriter(csvOutputPath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(results);
            }
        }
    }
}