using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions
{
    public class MissingExcelFieldAttributeException : FieldException
    {
        public MissingExcelFieldAttributeException() : base(ExceptionMessages.MissingExcelFieldAttribute)
        {
        }
    }
}
