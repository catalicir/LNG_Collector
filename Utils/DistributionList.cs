using Newtonsoft.Json;

namespace _LNG_Collector.Utils
{
    public class DistributionList
    {
        [JsonProperty("nume")]
        public string Name { get; set; }

        [JsonProperty("email")]
        public string[] Email { get; set; }

        [JsonProperty("files")]
        public string[] Files { get; set; }
    }
}