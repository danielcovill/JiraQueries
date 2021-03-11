namespace work_charts.Models.Jira
{
    public class JiraGroupRequest
    {
        public string groupname { get; set; }
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public bool includeinactiveusers { get; set; }

        public JiraGroupRequest(string _groupname)
        {
            groupname = _groupname;
            startAt = 0;
            maxResults = 100;
            includeinactiveusers = false;
        }
    }
}