using Newtonsoft.Json;

namespace _LNG_Collector.Utils
{
    public class DistributionList
    {
        [JsonProperty("email")]
        public string[] Email { get; set; }

        [JsonProperty("files")]
        public string[] Files { get; set; }
    }
}