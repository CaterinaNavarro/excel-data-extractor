using ExcelDataExtractor.Core.Enums;

namespace ExcelDataExtractor.Core.Interfaces
{
    public interface IExcelField
    {
        public string ColumnName { get; } 
        public bool Required { get; }
        public DataTypes? Type { get; }
    }
}
