using ExcelDataExtractor.Core.Enums;
using ExcelDataExtractor.Core.Interfaces;

namespace ExcelDataExtractor.Core.Models;

public class ExcelField : IExcelField
{
    public string ColumnName { get; set; } = null!;
    public bool Required { get; set; }
    public DataTypes? Type { get; set; }
}
