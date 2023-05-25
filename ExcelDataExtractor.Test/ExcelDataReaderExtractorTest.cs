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
            _requiredFieldMissingContent,
            _columnsNamesOnlyOnFirstSheetContent,
            _thirdSheetHasValuesContent;

        public ExcelDataReaderExtractorTest()
        {
            string filesRootPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files");
            string columnsWithDataPath = Path.Combine(filesRootPath, "ColumnsWithData.xlsx"),
                dataOnTwoSheetsPath = Path.Combine(filesRootPath, "DataOnTwoSheets.xlsx"),
                dataWithMissingColumnsNamesPath = Path.Combine(filesRootPath, "DataWithMissingColumnsNames.xlsx"),
                emptyPath = Path.Combine(filesRootPath, "Empty.xlsx"),
                notSupportedFormatPath = Path.Combine(filesRootPath, "UnsupportedFileType.docx"),
                requiredFieldMissingPath = Path.Combine(filesRootPath, "RequiredFieldMissing.xlsx"),
                columnsNamesOnlyOnFirstSheetPath = Path.Combine(filesRootPath, "ColumnsNamesOnlyOnFirstSheet.xlsx"),
                thirdSheetHasValuesPath = Path.Combine(filesRootPath, "ThirdSheetHasValues.xlsx");
                

            _excelDataReaderExtractor = new ExcelDataReaderExtractor();
            _columnsWithDataContent = File.ReadAllBytes(columnsWithDataPath);
            _dataOnTwoSheetsContent = File.ReadAllBytes(dataOnTwoSheetsPath);
            _dataWithMissingColumnsNamesContent = File.ReadAllBytes(dataWithMissingColumnsNamesPath);
            _emptyContent = File.ReadAllBytes(emptyPath);
            _unsupportedFileTypeContent = File.ReadAllBytes(notSupportedFormatPath);
            _requiredFieldMissingContent = File.ReadAllBytes(requiredFieldMissingPath);
            _columnsNamesOnlyOnFirstSheetContent = File.ReadAllBytes(columnsNamesOnlyOnFirstSheetPath);
            _thirdSheetHasValuesContent = File.ReadAllBytes(thirdSheetHasValuesPath);   
        }

        [Fact]
        public void Extract_All_Data_No_Convert_Model()
        {
            IEnumerable<IEnumerable<Dictionary<string, object?>>> excelData;

            excelData = _excelDataReaderExtractor.ProcessExtractData(_thirdSheetHasValuesContent, excludeSheetsWithNoneOrOneRows: false);

            Assert.True(excelData.Count() == 3 && excelData.Last().Count() == 1);
        }

        [Fact]
        public void Extract_Data_Excluding_Sheets_With_None_One_Row()
        {
            IEnumerable<IEnumerable<Dictionary<string, object?>>> excelData;

            excelData = _excelDataReaderExtractor.ProcessExtractData(_thirdSheetHasValuesContent, excludeSheetsWithNoneOrOneRows: true);

            Assert.True(excelData.Count() == 1 && excelData.First().Count() == 1);
        }

        [Fact]
        public void Extract_Data_Validate_Fields_No_Convert_Model()
        {
            IEnumerable<IEnumerable<Dictionary<string, object?>>> excelData;
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

            excelData = _excelDataReaderExtractor.ProcessExtractData(_dataOnTwoSheetsContent, fields, ignoreUnindicatedFields: true, excludeSheetsWithNoneOrOneRows: false);

            Assert.True(excelData.Count() == 2 && 
                excelData.First().Any(firstSheet => (int)firstSheet["FirstColumnNumber"]! == firstColumnFirstSheetValue) &&
                excelData.Last().Any(secondSheet => secondSheet["SecondColumnStringSecondSheet"]!.ToString() == secondColumnSecondSheetValue));
        }

        [Fact]
        public void Extract_Data_Sheet_Fields_Convert_Model()
        {
            IEnumerable<ExcelDataRow> excelDataSheet;
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

            Assert.True(excelDataSheet.Count() == 2);
        }

        [Fact]
        public void Extract_Data_Second_Sheet_Fields_Convert_Model()
        {
            IEnumerable<ExcelDataRowSecondSheet> excelDataSheet;
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
        public void Extract_Data_Sheet_Convert_Model_With_Fields_Attribute()
        {
            IEnumerable<ExcelDataRowWithFieldAttribute> excelDataSheet;

            excelDataSheet = _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRowWithFieldAttribute>(_columnsWithDataContent, ignoreUnindicatedFields: true);

            Assert.NotEmpty(excelDataSheet);
        }

        [Fact]
        public void Throw_Exception_Unsupported_File()
        {
            List<List<Dictionary<string, object?>>> excelData = new();

            Assert.Throws<UnsupportedFileException>(() => _excelDataReaderExtractor.ProcessExtractData(_unsupportedFileTypeContent, excludeSheetsWithNoneOrOneRows: true));
        }

        [Fact]
        public void Throw_Exception_Sheet_Has_Only_One_Row()
        {
            Assert.Throws<SheetHasOnlyOneRowException>(() => _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRowWithFieldAttribute>(_columnsNamesOnlyOnFirstSheetContent, ignoreUnindicatedFields: false));
        }

        [Fact]
        public void Throw_Exception_Sheet_Has_No_Row()
        {
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

            int sheetIndex = 1;

            Assert.Throws<SheetHasNoRowException>(() => _excelDataReaderExtractor.ProcessExtractDataSheet<ExcelDataRow>(_thirdSheetHasValuesContent, fields: fields, ignoreUnindicatedFields: false, sheetIndex: sheetIndex));
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

            Assert.Throws<SheetIndexNoExists>(() => _excelDataReaderExtractor.ProcessExtractData(_dataOnTwoSheetsContent, fields, ignoreUnindicatedFields: true, excludeSheetsWithNoneOrOneRows: false));
        }

        [Fact]
        public void Throw_Exception_File_Has_No_Row_Or_Data()
        {
            List<List<Dictionary<string, object?>>> excelData = new();

            Assert.Throws<FileHasNoDataException>(() => _excelDataReaderExtractor.ProcessExtractData(_emptyContent, excludeSheetsWithNoneOrOneRows: true));
        }

        [Fact]
        public void Throw_Exception_Values_Without_Column_Names()
        {
            List<List<Dictionary<string, object?>>> excelData = new();

            Assert.Throws<FieldHasValueNoColumnNameException>(() => _excelDataReaderExtractor.ProcessExtractData(_dataWithMissingColumnsNamesContent, excludeSheetsWithNoneOrOneRows: false));
        }

        [Fact]
        public void Throw_Exception_Convert_Model_Without_Fields_Attribute()
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
                    Required = false,
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
        public void Throw_Exception_Not_Indicated_Column_Name()
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

            Assert.Throws<NotIndicatedColumnNameException>(() => _excelDataReaderExtractor.ProcessExtractData(_columnsWithDataContent, fields, ignoreUnindicatedFields: false, excludeSheetsWithNoneOrOneRows: false));
        }

    }
}