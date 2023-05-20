using ExcelDataExtractor.Core.Enums;
using ExcelDataExtractor.Core.Interfaces;
using Newtonsoft.Json;

namespace ExcelDataExtractor.Core.Attributes
{
    /// <summary>
    /// Provides information of the field, implements <c>JsonPropertyAttribute</c> for matching the column name.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelFieldAttribute : Attribute, IExcelField
    {
        private readonly JsonPropertyAttribute _jsonPropertyAttribute;
        /// <summary>
        /// Column name of the sheet.
        /// </summary>
        public string ColumnName { get; } 
        public bool Required { get; }
        public DataTypes? Type { get; }

        public ExcelFieldAttribute(string columnName, bool required, DataTypes type)
        {
            _jsonPropertyAttribute = new JsonPropertyAttribute(columnName);
            ColumnName = _jsonPropertyAttribute.PropertyName!;
            Required = required; 
            Type = type;
        }
    }
}
