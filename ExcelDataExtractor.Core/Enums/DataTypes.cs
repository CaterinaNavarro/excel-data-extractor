using System.ComponentModel;

namespace ExcelDataExtractor.Core.Enums;

public enum DataTypes
{
    [Description("Integer number")]
    Integer = 1,

    [Description("Text")]
    String = 2,

    [Description("Date")]
    DateTime = 3
}
