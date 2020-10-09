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
        private readonly string _fileName;
        public ExcelDataParser(string fileName)
        {
            _fileName = fileName;
        }


        public DataTable ExcelParser(string sheetName, bool useHeaderRow)
        {
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
