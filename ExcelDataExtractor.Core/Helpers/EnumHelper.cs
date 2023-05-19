namespace ExcelDataExtractor.Core.Helpers;

internal static class EnumHelper
{
    internal static IEnumerable<T> GetValues<T>() where T : Enum
        => Enum.GetValues(typeof(T)).Cast<T>();
}
