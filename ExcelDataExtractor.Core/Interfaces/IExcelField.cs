using ExcelDataExtractor.Core.Enums;

namespace ExcelDataExtractor.Core.Interfaces
{
    public interface IExcelField
    {
        /// <summary>
        /// Column name of the sheet
        /// </summary>
        public string ColumnName { get; }

        /// <summary>
        /// If true it must have any value distinct from
        /// null or empty
        /// </summary>
        public bool Required { get; }

        /// <summary>
        /// Validate if value is in the correct type
        /// </summary>
        public DataTypes? Type { get; }
    }
}
