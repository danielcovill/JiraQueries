namespace work_charts
{
    public class JqlSearchRequest
    {
        public string query { get; set; }
        public int startAt { get; set; }
        public int maxResults { get; set; }
        public string[] fields { get; set; }
        public string validateQuery { get; set; }
        public string[] expand { get; set; }
        public string[] properties { get; set; }
        public bool fieldsByKeys { get; set; }
        public JqlSearchRequest(string query,
            string[] fields,
            string validateQuery = "strict",
            string[] expand = null,
            string[] properties = null,
            bool fieldsByKeys = false,
            int startAt = 0,
            int maxResults = 300)
        {
            this.query = query;
            this.startAt = startAt;
            this.maxResults = maxResults;
            this.fields = fields;
            this.validateQuery = validateQuery;
            this.expand = expand;
            this.fieldsByKeys = fieldsByKeys;
            this.properties = properties;
        }
    }
}