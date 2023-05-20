using System.ComponentModel;

namespace ExcelDataExtractor.Core.Enums;

/// <summary>
/// Field data types.
/// </summary>
public enum DataTypes
{
    [Description("Integer number")]
    Integer = 1,

    [Description("Text")]
    String = 2
}
