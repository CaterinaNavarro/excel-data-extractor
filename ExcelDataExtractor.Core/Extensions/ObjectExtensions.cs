using System.Reflection;

namespace ExcelDataExtractor.Core.Extensions;

public static class ObjectExtensions
{
    /// <summary>
    /// <c> Gets the specified attribute listing </c>
    /// </summary>
    /// <param name="value"> <c>Object value</c> </param>
    /// <returns> <c> List of attributes with their values </c> </returns>
    public static IEnumerable<T> GetAttributes<T>(this object value) where T : Attribute
    {
        if (value == null) throw new ArgumentNullException(nameof(value), "Value cannot be null");

        Type type = value.GetType();

        IEnumerable<object> allAttributes = type.IsEnum ?
            type.GetField(value.ToString()!)!.GetCustomAttributes(typeof(T), inherit: false) :
            type == typeof(PropertyInfo) || typeof(PropertyInfo).IsAssignableFrom(type) ? ((PropertyInfo)value).GetCustomAttributes(typeof(T), inherit: false) :
            type.GetCustomAttributes(typeof(T), inherit: true);

        IEnumerable<T> attributes = allAttributes.Cast<T>();

        return attributes;
    }

    /// <summary>
    /// <c> Gets a specified attribute </c>
    /// </summary>
    /// <param name="value"> <c>Object value</c> </param>
    /// <returns> <c> An attribute, if the value has more than one of the same type returns the first, 
    /// or null if it is not found </c> </returns>
    public static T? GetAttribute<T>(this object value) where T : Attribute
        => value.GetAttributes<T>().FirstOrDefault();
}

