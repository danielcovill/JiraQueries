namespace work_charts
{
    public class JqlSearchRequest
    {
        public string jql { get; set; }
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public string[] fields { get; set; }
        public string validateQuery { get; set; }
        public string[] expand { get; set; }
        public string[] properties { get; set; }
        public bool fieldsByKeys { get; set; }
        public JqlSearchRequest(string query,
            string[] fields = null,
            string validateQuery = "strict",
            string[] expand = null,
            string[] properties = null,
            bool fieldsByKeys = false,
            int startAt = 0,
            int maxResults = 100)
        {
            this.jql = query;
            this.startAt = startAt;
            this.maxResults = maxResults;
            this.fields = fields ?? new string[]
            {
                "key"
                , "customfield_10014" // Epic
                , "summary"
                , "assignee"
                , "customfield_10026" // Story Points
                , "issuetype"
                , "status"
                , "resolution"
                , "created"
                , "resolved"
                , "components"
                , "creator"
                , "resolutiondate"
            };
            this.validateQuery = validateQuery;
            this.expand = expand;
            this.fieldsByKeys = fieldsByKeys;
            this.properties = properties;
        }
    }
}