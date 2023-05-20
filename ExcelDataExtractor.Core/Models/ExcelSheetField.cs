using ExcelDataExtractor.Core.Enums;
using ExcelDataExtractor.Core.Interfaces;

namespace ExcelDataExtractor.Core.Models
{
    /// <summary>
    /// Indicate a field of a sheet
    /// </summary>
    public class ExcelSheetField : IExcelField
    {
        public string ColumnName { get; set; } = null!;
        public bool Required { get; set; }
        public DataTypes? Type { get; set; }

        /// <summary>
        /// Index Sheet
        /// </summary>
        public int SheetIndex { get; set; }
    }
}
