using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions
{
    public class NoColumnNameFirstRowException : ColumnException
    {
        public NoColumnNameFirstRowException() : base(ExceptionMessages.NoColumnNameFirstRow)
        {
        }
    }
}
