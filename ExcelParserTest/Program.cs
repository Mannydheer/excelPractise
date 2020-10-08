using ExcelDataReader;
using ExcelParserTest.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelParserTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

     

            //Using the DataSet Methods.
            ExcelFileParserWithAsDataSetMethod(@"C:\work\Ai-Projects\2.1\ReturnOnAssetsTTM.xlsx");

            //Using the Reader Methods. 
            //ExcelFileParserWithReaderMethod(@"C:\work\Ai-Projects\2.1\ReturnOnAssetsTTM.xlsx");

        }

        public static void ExcelFileParserWithReaderMethod(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                            //Most likely a switch case to handle the different sheets, which calls the appropriate methods. 
                            switch (reader.Name)
                            {
                              /*  case "Description":
                                    HandleDescriptionFactor(reader);
                                    break;
                                case "Timing":
                                    HandleTimingFactor(reader);
                                    break;*/
                                case "Ranking":
                                    HandleRankingFactor(reader);
                                    break;
                                default:
                                    break;
                            }                       
                     
                        //goes to the next sheet. 
                    } while (reader.NextResult());
                }
            }
        }

        public static void HandleDescriptionFactor(IExcelDataReader reader)
        {
            var descriptionModel = new Description();
            //loop through each row.
            while (reader.Read())
            {
                //Loop through each column of the the current row. 
                for (int i = 0; i < reader.FieldCount; i+=2)
                {

                    //gets the current value.
                    var currentCellValue = reader.IsDBNull(i) ? reader.GetValue(i) : reader.GetValue(i).ToString().Trim();

                    Console.WriteLine(currentCellValue);
                }
            }
        }
        public static void HandleTimingFactor(IExcelDataReader reader)
        {
            while (reader.Read())
            {
                    //Loop through each column of the the current row. 
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    if (reader.IsDBNull(i))
                        break;
                    //gets the current value.
                    var currentCellValue = reader.GetValue(i).ToString().Trim();

                    Console.WriteLine(currentCellValue);
                }

            }
        }

        public static void HandleRankingFactor (IExcelDataReader reader)
        {
            while (reader.Read())
            {
                //Loop through each column of the the current row. 
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    //gets the current value.
                    var currentCellValue = reader.GetValue(i);

                    Console.WriteLine(currentCellValue);
                }
            }

        }



















        //OTHER WAY TO PARSE EXCEL.
        public static void ExcelFileParserWithAsDataSetMethod(string filePath)
        {

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
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
                            UseHeaderRow = false,
                        }
                    });

                    var allDescription = new List<Description>();
                    //Get all tables.
                    var allTables = dataSet.Tables;

                    var sheetName = "Description";
                    //Get the table you want (based on sheet).
                    var dataTable = allTables[sheetName];

                    switch (sheetName)
                    {
                        case "Description":
                            DescriptionConverter(dataTable);
                            break;
                        case "Ranking_ExtraData":
                            RankingExtraDataConverter(dataTable);
                            break;
                        case "Ranking":
                            
                            break;
                        default:
                            break;
                    }


                }

            }

        }

        public static void DescriptionConverter(DataTable dataTable)
        {
            var convertDescriptionList = new List<KeyValuePair<string, string>>();


            //Loop through each row.
            foreach (DataRow itemsInRow in dataTable.Rows)
            {
                //Only in the case of Period since there are 3 column values. 
                //Check if it's empty/null 
                if (itemsInRow.ItemArray[0].ToString() == "Period")
                {
                    convertDescriptionList.Add(new KeyValuePair<string, string>("PeriodStartDate", itemsInRow.ItemArray[1].ToString()));
                    convertDescriptionList.Add(new KeyValuePair<string, string>("PeriodEndDate", itemsInRow.ItemArray[2].ToString()));
                }
                else
                {
                    convertDescriptionList.Add(new KeyValuePair<string, string>(itemsInRow.ItemArray[0].ToString(), itemsInRow.ItemArray[1].ToString()));
                }
            }

            var descriptionLists = new List<Description>();
            if (convertDescriptionList.Count > 0)
            {
                descriptionLists.Add(new Description {
                    Return = convertDescriptionList[0].Value,
                    PeriodStartDate = DateTime.Parse(convertDescriptionList[1].Value),
                    PeriodEndDate = DateTime.Parse(convertDescriptionList[2].Value),
                    RebalanceFrequency = convertDescriptionList[3].Value,
                    RankingMethod = convertDescriptionList[4].Value,
                    Slippage = int.Parse(convertDescriptionList[5].Value),
                    
                   
                });
            }
            Console.WriteLine(convertDescriptionList[0].Key);

        }

        public static void RankingExtraDataConverter(DataTable dataTable)
        {
            var rankingExtraDataLists = new List<RankingPerformance>();
            
            foreach (DataRow itemsInRow in dataTable.Rows)
            {
                var rankingList = new RankingPerformance()
                {               
                    Direction = itemsInRow["Direction"].ToString(),
                    Formula = itemsInRow["Formula"].ToString(),
                    Min = Convert.ToInt32(itemsInRow["Min"]),
                    Median = Convert.ToInt32(itemsInRow["Median"]),
                    Mean = Convert.ToInt32(itemsInRow["Mean"]),                    
                    //etc...
                };

                rankingExtraDataLists.Add(rankingList);
                     
            }
        }


        
    }
}
