using ExcelDataExtractor.Core.Extensions;
using Xunit;
using ExcelDataExtractor.Core.Enums;
using System.ComponentModel;
using Newtonsoft.Json;

namespace ExcelDataExtractor.Test.Helpers
{
    public class ObjectExtensionsTests
    {
        [Fact]
        public void Object_Has_The_Specific_Attribute()
        {
            DataTypes dataType = DataTypes.String;

            DescriptionAttribute? descriptionAttribute = dataType.GetAttribute<DescriptionAttribute>();

            Assert.NotNull(descriptionAttribute);
        }

        [Fact]
        public void Object_Has_Not_The_Specific_Attribute()
        {
            DataTypes dataType = DataTypes.String;

            JsonPropertyAttribute? jsonPropertyAttribute = dataType.GetAttribute<JsonPropertyAttribute>();

            Assert.Null(jsonPropertyAttribute);
        }
    }
}
