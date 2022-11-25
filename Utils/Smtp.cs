using Newtonsoft.Json;

namespace _LNG_Collector.Utils
{
    public class Smtp
    {
        [JsonProperty("server")]
        public string Server { get; set; }

        [JsonProperty("user")]
        public string User { get; set; }

        [JsonProperty("pwd")]
        public string Pwd { get; set; }
    }
}