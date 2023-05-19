using ExcelDataExtractor.Core.Interfaces;
using Xunit;
using ExcelDataExtractor.Core;
using ExcelDataExtractor.Test.Models;
using ExcelDataExtractor.Core.Models;
using ExcelDataExtractor.Core.Enums;
using ExcelDataExtractor.Core.Exceptions;

namespace ExcelDataExtractor.Test
{
    public class ExcelDataReaderExtractorTest
    {
        private readonly IExcelDataReaderExtractor _excelDataReaderExtractor;
        private readonly byte[] _columnsWithDataContent,
            _dataWithMissingColumnsNamesContent,
            _emptyContent,
            _unsupportedFileTypeContent,
            _requiredFieldMissingContent;

        public ExcelDataReaderExtractorTest()
        {
            string filesRootPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files");
            string columnsWithDataPath = Path.Combine(filesRootPath, "ColumnsWithData.xlsx"),
                dataWithMissingColumnsNamesPath = Path.Combine(filesRootPath, "DataWithMissingColumnsNames.xlsx"),
                emptyPath = Path.Combine(filesRootPath, "Empty.xlsx"),
                notSupportedFormatPath = Path.Combine(filesRootPath, "UnsupportedFileType.docx"),
                requiredFieldMissingPath = Path.Combine(filesRootPath, "RequiredFieldMissing.xlsx");
                

            _excelDataReaderExtractor = new ExcelDataReaderExtractor();
            _columnsWithDataContent = File.ReadAllBytes(columnsWithDataPath);
            _dataWithMissingColumnsNamesContent = File.ReadAllBytes(dataWithMissingColumnsNamesPath);
            _emptyContent = File.ReadAllBytes(emptyPath);
            _unsupportedFileTypeContent = File.ReadAllBytes(notSupportedFormatPath);
            _requiredFieldMissingContent = File.ReadAllBytes(requiredFieldMissingPath);
        }

        [Fact]
        public void Extract_Data_No_Parse_Model()
        {
            List<List<Dictionary<string, object?>>> excelData;

            excelData = _excelDataReaderExtractor.ProcessExtractData(_columnsWithDataContent);

            Assert.NotEmpty(excelData);
        }

        [Fact]
        public void Extract_Data_Sheet_Parse_Model()
        {
            List<ExcelDataRow> excelDataSheet;

            excelDataSheet = _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRow>(_columnsWithDataContent);
                
            Assert.NotEmpty(excelDataSheet);
        }

        [Fact]
        public void Extract_Data_Sheet_Fields_Parse_Model()
        {
            List<ExcelDataRow> excelDataSheet;
            List<ExcelField> fields = new()
            {
                new()
                {
                    ColumnName = "FirstColumnNumber",
                    Required = true,
                    Type = DataTypes.Integer
                }
            };

            excelDataSheet = _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRow>(_columnsWithDataContent, fields: fields, ignoreUnindicatedFields: true);

            Assert.NotEmpty(excelDataSheet);
        }

        [Fact]
        public void Extract_Data_Sheet_Parse_Model_With_Fields_Attribute()
        {
            List<ExcelDataRowWithFieldAttribute> excelDataSheet;

            excelDataSheet = _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRowWithFieldAttribute>(_columnsWithDataContent, ignoreUnindicatedFields: true);

            Assert.NotEmpty(excelDataSheet);
        }

        [Fact]
        public void Throw_Exception_Unsupported_File()
        {
            List<List<Dictionary<string, object?>>> excelData = new();

            Assert.Throws<UnsupportedFileException>(() => _excelDataReaderExtractor.ProcessExtractData(_unsupportedFileTypeContent));
        }

        [Fact]
        public void Throw_Exception_Sheet_Index_No_Exists()
        {
            int index = 5;

            Assert.Throws<UnsupportedFileException>(() => _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRow>(_unsupportedFileTypeContent, index));
        }

        [Fact]
        public void Throw_Exception_Empty_File()
        {
            List<List<Dictionary<string, object?>>> excelData = new();

            Assert.Throws<EmptySheetException>(() => _excelDataReaderExtractor.ProcessExtractData(_emptyContent));
        }

        [Fact]
        public void Throw_Exception_Values_Without_Column_Names()
        {
            List<List<Dictionary<string, object?>>> excelData = new();

            Assert.Throws<FieldHasValueNoColumnNameException>(() => _excelDataReaderExtractor.ProcessExtractData(_dataWithMissingColumnsNamesContent));
        }

        [Fact]
        public void Throw_Exception_Parse_Model_Without_Fields_Attribute()
        {
            Assert.Throws<MissingExcelFieldAttributeException>(() => _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRow>(_columnsWithDataContent, ignoreUnindicatedFields: false));
        }

        [Fact]
        public void Throw_Exception_Data_Type_Field_Incorrect()
        {
            List<ExcelField> fields = new()
            {
                new()
                {
                    ColumnName = "SecondColumnValue",
                    Required = true,
                    Type = DataTypes.Integer
                }
            };

            Assert.Throws<ColumnValueNotMatchFieldDataTypeException>(() => _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRow>(_columnsWithDataContent, fields, ignoreUnindicatedFields: true));
        }

        [Fact]
        public void Throw_Exception_Required_Field()
        {
            Assert.Throws<RequiredFieldException>(() => _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRowWithFieldAttribute>(_requiredFieldMissingContent, ignoreUnindicatedFields: true));
        }
    }
}