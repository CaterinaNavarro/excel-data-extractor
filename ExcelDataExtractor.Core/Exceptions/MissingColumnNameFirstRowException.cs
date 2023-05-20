using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when a sheet does not contain any column name in the first row.
/// </summary>
public class MissingColumnNameFirstRowException : ColumnException
{
    public MissingColumnNameFirstRowException() : base(ExceptionMessages.MissingColumnNameFirstRow)
    {
    }
}
