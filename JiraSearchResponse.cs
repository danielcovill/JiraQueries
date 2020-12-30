using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace work_charts
{
    public class Issuetype    {
        public string self { get; set; } 
        public string id { get; set; } 
        public string description { get; set; } 
        public string iconUrl { get; set; } 
        public string name { get; set; } 
        public bool subtask { get; set; } 
        public int avatarId { get; set; } 
    }

    public class AvatarUrls    {
        [JsonPropertyName("48x48")]
        public string _48x48 { get; set; } 
        [JsonPropertyName("24x24")]
        public string _24x24 { get; set; } 
        [JsonPropertyName("16x16")]
        public string _16x16 { get; set; } 
        [JsonPropertyName("32x32")]
        public string _32x32 { get; set; } 
    }

    public class Project    {
        public string self { get; set; } 
        public string id { get; set; } 
        public string key { get; set; } 
        public string name { get; set; } 
        public string projectTypeKey { get; set; } 
        public bool simplified { get; set; } 
        public AvatarUrls avatarUrls { get; set; } 
    }

    public class Regression {
        public string self { get; set; } 
        public string value { get; set; } 
        public string id { get; set; } 
    }

    public class FixVersion    {
        public string self { get; set; } 
        public string id { get; set; } 
        public string description { get; set; } 
        public string name { get; set; } 
        public bool archived { get; set; } 
        public bool released { get; set; } 
        public DateTime releaseDate { get; set; } 
    }

    public class Resolution    {
        public string self { get; set; } 
        public string id { get; set; } 
        public string description { get; set; } 
        public string name { get; set; } 
    }

    public class Customfield10027    {
        public string self { get; set; } 
        public string accountId { get; set; } 
        public AvatarUrls avatarUrls { get; set; } 
        public string displayName { get; set; } 
        public bool active { get; set; } 
        public string timeZone { get; set; } 
        public string accountType { get; set; } 
        public string emailAddress { get; set; } 
    }

    public class Watches    {
        public string self { get; set; } 
        public int watchCount { get; set; } 
        public bool isWatching { get; set; } 
    }

    public class Issuerestrictions    {
    }

    public class Issuerestriction    {
        public Issuerestrictions issuerestrictions { get; set; } 
        public bool shouldDisplay { get; set; } 
    }

    public class Priority    {
        public string self { get; set; } 
        public string iconUrl { get; set; } 
        public string name { get; set; } 
        public string id { get; set; } 
    }

    public class NonEditableReason    {
        public string reason { get; set; } 
        public string message { get; set; } 
    }

    public class Customfield10018    {
        public bool hasEpicLinkFieldDependency { get; set; } 
        public bool showField { get; set; } 
        public NonEditableReason nonEditableReason { get; set; } 
    }

    public class Type    {
        public string id { get; set; } 
        public string name { get; set; } 
        public string inward { get; set; } 
        public string outward { get; set; } 
        public string self { get; set; } 
    }

    public class StatusCategory    {
        public string self { get; set; } 
        public int id { get; set; } 
        public string key { get; set; } 
        public string colorName { get; set; } 
        public string name { get; set; } 
    }

    public class Status    {
        public string self { get; set; } 
        public string description { get; set; } 
        public string iconUrl { get; set; } 
        public string name { get; set; } 
        public string id { get; set; } 
        public StatusCategory statusCategory { get; set; } 
    }

    public class InwardIssue    {
        public string id { get; set; } 
        public string key { get; set; } 
        public string self { get; set; } 
        public Fields fields { get; set; } 
    }

    public class OutwardIssue    {
        public string id { get; set; } 
        public string key { get; set; } 
        public string self { get; set; } 
        public Fields fields { get; set; } 
    }

    public class Issuelink    {
        public string id { get; set; } 
        public string self { get; set; } 
        public Type type { get; set; } 
        public InwardIssue inwardIssue { get; set; } 
        public OutwardIssue outwardIssue { get; set; } 
    }

    public class Assignee    {
        public string self { get; set; } 
        public string accountId { get; set; } 
        public AvatarUrls avatarUrls { get; set; } 
        public string displayName { get; set; } 
        public bool active { get; set; } 
        public string timeZone { get; set; } 
        public string accountType { get; set; } 
    }

    public class Component    {
        public string self { get; set; } 
        public string id { get; set; } 
        public string name { get; set; } 
        public string description { get; set; } 
    }

    public class Author    {
        public string self { get; set; } 
        public string accountId { get; set; } 
        public AvatarUrls avatarUrls { get; set; } 
        public string displayName { get; set; } 
        public bool active { get; set; } 
        public string timeZone { get; set; } 
        public string accountType { get; set; } 
    }

    public class Attachment    {
        public string self { get; set; } 
        public string id { get; set; } 
        public string filename { get; set; } 
        public Author author { get; set; } 
        public DateTime created { get; set; } 
        public int size { get; set; } 
        public string mimeType { get; set; } 
        public string content { get; set; } 
        public string thumbnail { get; set; } 
    }

    public class Creator    {
        public string self { get; set; } 
        public string accountId { get; set; } 
        public AvatarUrls avatarUrls { get; set; } 
        public string displayName { get; set; } 
        public bool active { get; set; } 
        public string timeZone { get; set; } 
        public string accountType { get; set; } 
        public string emailAddress { get; set; } 
    }

    public class Subtask    {
        public string id { get; set; } 
        public string key { get; set; } 
        public string self { get; set; } 
        public Fields fields { get; set; } 
    }

    public class Reporter    {
        public string self { get; set; } 
        public string accountId { get; set; } 
        public AvatarUrls avatarUrls { get; set; } 
        public string displayName { get; set; } 
        public bool active { get; set; } 
        public string timeZone { get; set; } 
        public string accountType { get; set; } 
        public string emailAddress { get; set; } 
    }

    public class Aggregateprogress    {
        public int progress { get; set; } 
        public int total { get; set; } 
    }

    public class Progress    {
        public int progress { get; set; } 
        public int total { get; set; } 
    }

    public class Votes    {
        public string self { get; set; } 
        public int votes { get; set; } 
        public bool hasVoted { get; set; } 
    }

    public class UpdateAuthor    {
        public string self { get; set; } 
        public string accountId { get; set; } 
        public AvatarUrls avatarUrls { get; set; } 
        public string displayName { get; set; } 
        public bool active { get; set; } 
        public string timeZone { get; set; } 
        public string accountType { get; set; } 
        public string emailAddress { get; set; } 
    }

    public class Comment    {
        public List<Comment> comments { get; set; } 
        public int maxResults { get; set; } 
        public int total { get; set; } 
        public int startAt { get; set; } 
    }

    public class Worklog    {
        public int startAt { get; set; } 
        public int maxResults { get; set; } 
        public int total { get; set; } 
        public List<object> worklogs { get; set; } 
    }

    public class Fields    {
        public DateTime statuscategorychangedate { get; set; } 
        public Issuetype issuetype { get; set; } 
        public object timespent { get; set; } 
        public object customfield_10030 { get; set; } 
        public Project project { get; set; } 
        [JsonPropertyName("customfield_10031")]
        public List<Regression> regressions { get; set; } 
        public List<FixVersion> fixVersions { get; set; } 
        public object aggregatetimespent { get; set; } 
        public Resolution resolution { get; set; } 
        public List<Customfield10027> customfield_10027 { get; set; } 
        public object customfield_10029 { get; set; } 
        public DateTime? resolutiondate { get; set; } 
        public int workratio { get; set; } 
        public Watches watches { get; set; } 
        public Issuerestriction issuerestriction { get; set; } 
        public DateTime? lastViewed { get; set; } 
        public DateTime created { get; set; } 
        public object customfield_10020 { get; set; } 
        public object customfield_10021 { get; set; } 
        public DateTime? customfield_10022 { get; set; } 
        public string customfield_10023 { get; set; } 
        public Priority priority { get; set; } 
        public string customfield_10024 { get; set; } 
        [JsonPropertyName("customfield_10026")]
        public double? storyPoints { get; set; } 
        public List<string> labels { get; set; } 
        public object customfield_10016 { get; set; } 
        public object customfield_10017 { get; set; } 
        public Customfield10018 customfield_10018 { get; set; } 
        public string customfield_10019 { get; set; } 
        public object aggregatetimeoriginalestimate { get; set; } 
        public object timeestimate { get; set; } 
        public List<object> versions { get; set; } 
        public List<Issuelink> issuelinks { get; set; } 
        public Assignee assignee { get; set; } 
        public DateTime updated { get; set; } 
        public Status status { get; set; } 
        public List<Component> components { get; set; } 
        public object timeoriginalestimate { get; set; } 
        public string description { get; set; } 
        public object customfield_10010 { get; set; } 
        [JsonPropertyName("customfield_10014")]
        public string epicLink { get; set; } 
        public object customfield_10015 { get; set; } 
        public object customfield_10005 { get; set; } 
        public object customfield_10006 { get; set; } 
        public object security { get; set; } 
        public object customfield_10007 { get; set; } 
        public object customfield_10008 { get; set; } 
        public object customfield_10009 { get; set; } 
        public object aggregatetimeestimate { get; set; } 
        public List<Attachment> attachment { get; set; } 
        public string summary { get; set; } 
        public Creator creator { get; set; } 
        public List<Subtask> subtasks { get; set; } 
        public Reporter reporter { get; set; } 
        public Aggregateprogress aggregateprogress { get; set; } 
        public string customfield_10000 { get; set; } 
        public object customfield_10001 { get; set; } 
        public object customfield_10002 { get; set; } 
        public object customfield_10003 { get; set; } 
        public object customfield_10004 { get; set; } 
        public string environment { get; set; } 
        public object duedate { get; set; } 
        public Progress progress { get; set; } 
        public Votes votes { get; set; } 
        public Comment comment { get; set; } 
        public Worklog worklog { get; set; } 
    }

    public class Issue    {
        public string expand { get; set; } 
        public string id { get; set; } 
        public string self { get; set; } 
        public string key { get; set; } 
        public Fields fields { get; set; } 
    }

    public class Names    {
        public string statuscategorychangedate { get; set; } 
        public string issuetype { get; set; } 
        public string timespent { get; set; } 
        public string customfield_10030 { get; set; } 
        public string project { get; set; } 
        [JsonPropertyName("customfield_10031")]
        public string regression { get; set; } 
        public string fixVersions { get; set; } 
        public string aggregatetimespent { get; set; } 
        public string resolution { get; set; } 
        public string customfield_10027 { get; set; } 
        public string customfield_10029 { get; set; } 
        public DateTime resolutiondate { get; set; } 
        public string workratio { get; set; } 
        public string watches { get; set; } 
        public string issuerestriction { get; set; } 
        public string lastViewed { get; set; } 
        public DateTime created { get; set; } 
        public string customfield_10020 { get; set; } 
        public string customfield_10021 { get; set; } 
        public string customfield_10022 { get; set; } 
        public string customfield_10023 { get; set; } 
        public string priority { get; set; } 
        public string customfield_10024 { get; set; } 
        [JsonPropertyName("customfield_10026")]
        public string storyPoints { get; set; } 
        public string labels { get; set; } 
        public string customfield_10016 { get; set; } 
        public string customfield_10017 { get; set; } 
        public string customfield_10018 { get; set; } 
        public string customfield_10019 { get; set; } 
        public string aggregatetimeoriginalestimate { get; set; } 
        public string timeestimate { get; set; } 
        public string versions { get; set; } 
        public string issuelinks { get; set; } 
        public string assignee { get; set; } 
        public DateTime updated { get; set; } 
        public string status { get; set; } 
        public string components { get; set; } 
        public string timeoriginalestimate { get; set; } 
        public string description { get; set; } 
        public string customfield_10010 { get; set; } 
        [JsonPropertyName("customfield_10014")]
        public string epicLink { get; set; } 
        public string customfield_10015 { get; set; } 
        public string customfield_10005 { get; set; } 
        public string customfield_10006 { get; set; } 
        public string security { get; set; } 
        public string customfield_10007 { get; set; } 
        public string customfield_10008 { get; set; } 
        public string customfield_10009 { get; set; } 
        public string aggregatetimeestimate { get; set; } 
        public string attachment { get; set; } 
        public string summary { get; set; } 
        public string creator { get; set; } 
        public string subtasks { get; set; } 
        public string reporter { get; set; } 
        public string aggregateprogress { get; set; } 
        public string customfield_10000 { get; set; } 
        public string customfield_10001 { get; set; } 
        public string customfield_10002 { get; set; } 
        public string customfield_10003 { get; set; } 
        public string customfield_10004 { get; set; } 
        public string environment { get; set; } 
        public DateTime duedate { get; set; } 
        public string progress { get; set; } 
        public string votes { get; set; } 
        public string comment { get; set; } 
        public string worklog { get; set; } 
    }

    public class JiraSearchResponse {
        public string expand { get; set; } 
        public int startAt { get; set; } 
        public int maxResults { get; set; } 
        public int total { get; set; } 
        public List<Issue> issues { get; set; } 
        public List<string> warningMessages { get; set; } 
        public Names names { get; set; } 
        public int GetTicketCount(string ticketType = null)
        { 
            if(ticketType == null) 
            {
                return total;
            }
            else
            {
                return issues.Where(issue => issue.fields.issuetype.name == ticketType).Count();
            }
        }
        public double GetPointTotal(string ticketType = null)
        { 
            if(ticketType == null) 
            {
                return issues.Sum(issue => issue.fields.storyPoints) ?? 0;
            }
            else 
            {
                return issues.Where(issue => issue.fields.issuetype.name == ticketType).Sum(issue => issue.fields.storyPoints) ?? 0;
            }
        }
    }
}