using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// Exception thrown when the sheet/s of the file are empty or have only one row.
/// </summary>
public class FileHasNoDataException : FileException
{
    public FileHasNoDataException() : base(ExceptionMessages.FileHasNoData)
    {
    }
}
