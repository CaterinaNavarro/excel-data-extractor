using Newtonsoft.Json;

namespace ExcelDataExtractor.Test.Models
{
    internal class ExcelDataRowSecondSheet
    {
        [JsonProperty("FirstColumnNumberSecondSheet")]
        public int FirstColumn { get; set; }

        [JsonProperty("SecondColumnStringSecondSheet")]
        public string SecondColumn { get; set; } = null!;
    }
}
