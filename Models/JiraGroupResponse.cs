using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace work_charts
{
    public class User
    {
        public string self { get; set; } 
        public string name { get; set; } 
        public string key { get; set; } 
        public string accountId { get; set; } 
        public string emailAddress { get; set; } 
        public AvatarUrls avatarUrls { get; set; } 
        public string displayName { get; set; } 
        public bool active { get; set; } 
        public string timeZone { get; set; } 
        public string accountType { get; set; } 
    }

    public class JiraGroupResponse
    {
        public string self { get; set; } 
        public string nextPage { get; set; } 
        public int maxResults { get; set; } 
        public int startAt { get; set; } 
        public int total { get; set; } 
        public bool isLast { get; set; } 
        [JsonPropertyName("values")]
        public List<User> users { get; set; } 
    }
}