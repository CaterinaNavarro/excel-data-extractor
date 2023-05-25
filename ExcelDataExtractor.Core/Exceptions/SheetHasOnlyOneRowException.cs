using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when a sheet has only one row.
/// </summary>
public class SheetHasOnlyOneRowException : ColumnException
{
    public SheetHasOnlyOneRowException(string message) : base(message)
    {
    }
}
