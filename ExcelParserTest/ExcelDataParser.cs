using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace ExcelParserTest
{
    public class ExcelDataParser
    {
        private string _fileName;
        public ExcelDataParser()
        {
            _fileName = @"C:\work\Ai-Projects\2.1\ReturnOnAssetsTTM.xlsx";
        }


        public DataTable ExcelParser(string sheetName, bool useHeaderRow)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(_fileName, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    //Option to specify configurations. 
                    var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = useHeaderRow,
                        }
                    });

                    var allTables = dataSet.Tables;

                    //Get the table you want (based on sheet).
                    var dataTable = allTables[sheetName];

                    return dataTable;
                }

            }
        }


    }
}
