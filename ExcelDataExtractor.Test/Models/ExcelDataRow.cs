using Newtonsoft.Json;

namespace ExcelDataExtractor.Test.Models
{
    internal class ExcelDataRow
    {
        [JsonProperty("FirstColumnNumber")]
        public int FirstColumn { get; set; }

        [JsonProperty("SecondColumnValue")]
        public string SecondColumn { get; set; } = null!;
    }
}
