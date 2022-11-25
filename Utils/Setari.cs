using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace _LNG_Collector.Utils
{
    internal class Setari
    {
        [JsonProperty("input_files")]
        public IEnumerable<string> InputFiles { get; set; }

        [JsonProperty("distribution_list")]
        public IEnumerable<DistributionList> DistributionList { get; set; }

        [JsonProperty("smtp")]
        public Smtp Smtp { get; set; }
    }
}
