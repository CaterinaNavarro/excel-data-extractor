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
            _dataOnTwoSheetsContent,
            _dataWithMissingColumnsNamesContent,
            _emptyContent,
            _unsupportedFileTypeContent,
            _requiredFieldMissingContent;

        public ExcelDataReaderExtractorTest()
        {
            string filesRootPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files");
            string columnsWithDataPath = Path.Combine(filesRootPath, "ColumnsWithData.xlsx"),
                dataOnTwoSheetsPath = Path.Combine(filesRootPath, "DataOnTwoSheets.xlsx"),
                dataWithMissingColumnsNamesPath = Path.Combine(filesRootPath, "DataWithMissingColumnsNames.xlsx"),
                emptyPath = Path.Combine(filesRootPath, "Empty.xlsx"),
                notSupportedFormatPath = Path.Combine(filesRootPath, "UnsupportedFileType.docx"),
                requiredFieldMissingPath = Path.Combine(filesRootPath, "RequiredFieldMissing.xlsx");
                

            _excelDataReaderExtractor = new ExcelDataReaderExtractor();
            _columnsWithDataContent = File.ReadAllBytes(columnsWithDataPath);
            _dataOnTwoSheetsContent = File.ReadAllBytes(dataOnTwoSheetsPath);
            _dataWithMissingColumnsNamesContent = File.ReadAllBytes(dataWithMissingColumnsNamesPath);
            _emptyContent = File.ReadAllBytes(emptyPath);
            _unsupportedFileTypeContent = File.ReadAllBytes(notSupportedFormatPath);
            _requiredFieldMissingContent = File.ReadAllBytes(requiredFieldMissingPath);
        }

        [Fact]
        public void Extract_All_Data_No_Parse_Model()
        {
            List<List<Dictionary<string, object?>>> excelData;

            excelData = _excelDataReaderExtractor.ProcessExtractData(_columnsWithDataContent);

            Assert.NotEmpty(excelData);
        }

        [Fact]
        public void Extract_Data_Validate_Fields_No_Parse_Model()
        {
            List<List<Dictionary<string, object?>>> excelData;
            List<ExcelSheetField> fields = new()
            {
                new()
                {
                    ColumnName = "FirstColumnNumber",
                    Required = true,
                    Type = DataTypes.Integer,
                    SheetIndex = 0,
                },
                new()
                {
                    ColumnName = "SecondColumnStringSecondSheet",
                    Required = true,
                    Type= DataTypes.String,
                    SheetIndex = 1
                }
            };

            int firstColumnFirstSheetValue = 1;
            string secondColumnSecondSheetValue = "fifth value";

            excelData = _excelDataReaderExtractor.ProcessExtractData(_dataOnTwoSheetsContent, fields, ignoreUnindicatedFields: true);

            Assert.True(excelData[0].Any(firstSheet => (int)firstSheet["FirstColumnNumber"]! == firstColumnFirstSheetValue) &&
                        excelData[1].Any(secondSheet => secondSheet["SecondColumnStringSecondSheet"]!.ToString() == secondColumnSecondSheetValue));
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
                },
                new ()
                {
                    ColumnName = "SecondColumnString",
                    Required = true,
                    Type = DataTypes.String
                }
            };

            excelDataSheet = _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRow>(_columnsWithDataContent, fields: fields, ignoreUnindicatedFields: true);

            Assert.NotEmpty(excelDataSheet);
        }

        [Fact]
        public void Extract_Data_Second_Sheet_Fields_Parse_Model()
        {
            List<ExcelDataRowSecondSheet> excelDataSheet;
            List<ExcelField> fields = new()
            {
                new()
                {
                    ColumnName = "FirstColumnNumberSecondSheet",
                    Required = true,
                    Type = DataTypes.Integer
                },
                new()
                {
                    ColumnName = "SecondColumnStringSecondSheet",
                    Required = false,
                    Type = DataTypes.String
                }
            };

            int firstColumnValue = 5;
            string secondColumnValue = "fifth value";

            excelDataSheet = _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRowSecondSheet>(_dataOnTwoSheetsContent, fields: fields, ignoreUnindicatedFields: false, sheetIndex: 1);

            Assert.Contains(excelDataSheet, x => x.FirstColumn == firstColumnValue && x.SecondColumn == secondColumnValue);
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
        public void Throw_Exception_Sheet_No_Exists()
        {
            List<ExcelSheetField> fields = new()
            {
                new()
                {
                    ColumnName = "FirstColumnNumber",
                    Required = true,
                    Type = DataTypes.Integer,
                    SheetIndex = 0,
                },
                new()
                {
                    ColumnName = "SecondColumnStringSecondSheet",
                    Required = true,
                    Type = DataTypes.String,
                    SheetIndex = 5
                }
            };

            Assert.Throws<SheetIndexNoExists>(() => _excelDataReaderExtractor.ProcessExtractData(_dataOnTwoSheetsContent, fields, ignoreUnindicatedFields: true));
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
                    ColumnName = "SecondColumnString",
                    Required = true,
                    Type = DataTypes.Integer
                }
            };

            Assert.Throws<FieldValueTypeDifferentFieldDataTypeException>(() => _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRow>(_columnsWithDataContent, fields, ignoreUnindicatedFields: true));
        }

        [Fact]
        public void Throw_Exception_Required_Field()
        {
            Assert.Throws<RequiredFieldException>(() => _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRowWithFieldAttribute>(_requiredFieldMissingContent, ignoreUnindicatedFields: true));
        }

        [Fact]
        public void Throw_Exception_Not_Indicated_Field()
        {
            List<ExcelSheetField> fields = new()
            {
                new()
                {
                    ColumnName = "FirstColumnNumber",
                    Required = true,
                    Type = DataTypes.Integer,
                    SheetIndex = 0
                }
            };

            Assert.Throws<NotIndicatedFieldException>(() => _excelDataReaderExtractor.ProcessExtractData(_columnsWithDataContent, fields, ignoreUnindicatedFields: false));
        }

    }
}