using ExcelDataExtractor.Core.Attributes;
using ExcelDataExtractor.Core.Enums;
using Newtonsoft.Json;

namespace ExcelDataExtractor.Test.Models
{
    internal class ExcelDataRowWithFieldAttribute
    {
        [ExcelField(columnName: "FirstColumnNumber", required: true, type: DataTypes.Integer)]
        public int FirstColumn { get; set; }
        
        [ExcelField(columnName: "SecondColumnString", required: true, type: DataTypes.String)]
        public string SecondColumn { get; set; } = null!;
    }
}
