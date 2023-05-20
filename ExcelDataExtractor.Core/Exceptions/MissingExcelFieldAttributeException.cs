using ExcelDataExtractor.Core.Constants;

namespace ExcelDataExtractor.Core.Exceptions;

/// <summary>
/// <para>Exception thrown when using the method to extract data and parse into an specific object type <c>T</c> without indicating
/// the list of fields apart.</para>
///  In this case, the type <c>T</c> must include the <c> ExcelFieldAttribute </c> in all of its properties.
/// </summary>
public class MissingExcelFieldAttributeException : FieldException
{
    public MissingExcelFieldAttributeException() : base(ExceptionMessages.MissingExcelFieldAttribute)
    {
    }
}
