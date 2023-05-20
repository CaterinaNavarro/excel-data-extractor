using ExcelDataExtractor.Core.Enums;

namespace ExcelDataExtractor.Core.Helpers;

internal class TypeConverterHelper
{
    private readonly Dictionary<DataTypes, Type> _converters = new()
    {
        { DataTypes.Integer, typeof(int) },
        { DataTypes.String, typeof(string) }
    };

    internal bool TryParse(object value, DataTypes dataType, out object? valueConverted)
    {
        bool converted = false;

        try
        {
            if (value is null)
                throw new NullReferenceException("Value cannot be null");

            if (!_converters.Any(x => x.Key == dataType))
                throw new NotImplementedException();

            Type type = _converters[dataType];

            valueConverted = Convert.ChangeType(value, type);

            converted = true;
        }
        catch (FormatException)
        {
            valueConverted = null;
        }

        return converted;
    }
}
