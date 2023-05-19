using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions
{
    public class ExcelFieldColumnNameNullEmptyException : FieldException
    {
        public ExcelFieldColumnNameNullEmptyException() : base(ExceptionMessages.ExcelFieldColumnNameNullEmpty)
        {
        }
    }
}
