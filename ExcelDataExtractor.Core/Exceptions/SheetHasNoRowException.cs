namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when the sheet has no rows.
/// </summary>
public class SheetHasNoRowException : SheetException
{
    public SheetHasNoRowException(string message) : base(message)
    {
    }
}
