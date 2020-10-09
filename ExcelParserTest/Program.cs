using ExcelDataReader;
using ExcelParserTest.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;

namespace ExcelParserTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);



            //Using the DataSet Methods.
            var excelParser = new ExcelDataParser(@"C:\work\Ai-Projects\2.1\ReturnOnAssetsTTM.xlsx");
            

            var description = DescriptionConverter(excelParser);
            var timing = TimingConverter(excelParser);
            var sectorPerformance = SectorPerformanceConverter(excelParser);
            var combinations = CombinationsConverter(excelParser);


            var factors = new Factors
            {
                Description = description,
                Timing = timing,
                SectorPerformance = sectorPerformance,
                Combinations = combinations
            };

            var factorsJsonData = JsonConverter(factors);

            //
            RankingPerformanceConverter(excelParser);
         
        }

    
        public static Description DescriptionConverter(ExcelDataParser excelParser)
        {
            //Pass the sheet name and if we want headerRow or not. 
            var dataTable = excelParser.ExcelParser("Description", false);

            var convertDescriptionList = new List<string>();

            //Loop through each row.
            foreach (DataRow itemsInRow in dataTable.Rows)
            {
                //Only in the case of Period since there are 3 column values. 
                //TODO Check if it's empty/null 
                if (itemsInRow.ItemArray[0].ToString() == "Period")
                {
                    convertDescriptionList.Add(itemsInRow.ItemArray[1].ToString());
                    convertDescriptionList.Add(itemsInRow.ItemArray[2].ToString());
                }
                else
                {
                    convertDescriptionList.Add(itemsInRow.ItemArray[1].ToString());
                }
            }

            Description descriptionObj;

            if (convertDescriptionList.Count > 0)
            {
                descriptionObj = new Description {
                    Return = convertDescriptionList[0],
                    PeriodStartDate = DateTime.Parse(convertDescriptionList[1]),
                    PeriodEndDate = DateTime.Parse(convertDescriptionList[2]),
                    RebalanceFrequency = convertDescriptionList[3],
                    RankingMethod = convertDescriptionList[4],
                    Slippage = Convert.ToDecimal((convertDescriptionList[5])),
                    TransactionType = convertDescriptionList[6]
                    //TODO
                    //MISSING PROPERTIES. 
                                      
                };

                return descriptionObj;

            }
            //refactor
            return null;

           

            

        }
        //Ranking_ExtraData & Ranking Sheet.
        public static void RankingPerformanceConverter(ExcelDataParser excelParser)
        {
            //Get the table.
            var dataTableForRankingExtraDataSheet = excelParser.ExcelParser("Ranking_ExtraData", true);

            
            RankingPerformance rankingPerformance;

            foreach (DataRow itemsInRow in dataTableForRankingExtraDataSheet.Rows)
            {
                  rankingPerformance = new RankingPerformance()
                {               
                    Direction = itemsInRow["Direction"].ToString(),
                    Formula = itemsInRow["Formula"].ToString(),
                    Min = Convert.ToDecimal(itemsInRow["Min"]),
                    Median = Convert.ToDecimal(itemsInRow["Median"]),
                    Mean = Convert.ToDecimal(itemsInRow["Mean"]),    
                    Max = Convert.ToDecimal(itemsInRow["Max"]),
                    First = Convert.ToDecimal(itemsInRow["First"]),
                    Last = Convert.ToDecimal(itemsInRow["Last"]),
                    Delta = Convert.ToDecimal(itemsInRow["Delta"]),
                    Slope = Convert.ToDecimal(itemsInRow["Slope"]),
                    StandardDeviation = Convert.ToDecimal(itemsInRow["StandardDeviation"])
                  };
                     
            }


            //Now need to isntantiate ranking dat ana dfill the fields. 
            var dataTableForRankingSheet = excelParser.ExcelParser("Ranking", true);

            List<object> rankingDataListsFirstRow;

            List<object> rankingDataListsSecondRow;

            foreach (DataRow itemsInRow in dataTableForRankingSheet.Rows)
            {
               rankingDataListsFirstRow = itemsInRow.ItemArray.ToList();              
            }

            foreach (DataRow itemsInRow in dataTableForRankingSheet.Rows)
            {
                rankingDataListsSecondRow = itemsInRow.ItemArray.ToList();
            }
        }

        //TODO
        //Performance Sheet. 

        //Contributions. 

        public static Combinations CombinationsConverter(ExcelDataParser excelParser)
        {

            var combinations = new Combinations();

            //Contributions Sheet
            var dataTable = excelParser.ExcelParser("Contributions", true);
            var contributionLists = new List<Contributions>();


            foreach (DataRow itemsInRow in dataTable.Rows)
            {
                contributionLists.Add(new Contributions
                {
                    NameEn = itemsInRow.ItemArray[0].ToString(),
                    NameFr = itemsInRow.ItemArray[1].ToString(),
                    Value = Convert.ToDecimal(itemsInRow.ItemArray[2])
                });
            }

            combinations.Contributions = contributionLists;

            //Correlations Sheet
            //TODO.
            var dataTableCorrelations = excelParser.ExcelParser("Correlation", false);


            foreach (DataRow itemsInRow in dataTableCorrelations.Rows)
            {
                var correlationList = new List<decimal>();

                for (int i = 0; i < itemsInRow.ItemArray.Count(); i++)
                {
                    //validate? Maybe for empty or spaces?
                    if(itemsInRow.ItemArray[i] != DBNull.Value)
                    {
                        correlationList.Add(Convert.ToDecimal(itemsInRow.ItemArray[i]));
                    }                                  
                }
                //add to combinations object which holds all rows.
                combinations.Correlations.Add(correlationList);
            }

            return combinations;
        }

        public static List<Timing> TimingConverter(ExcelDataParser excelParser)
        {  
            //TODO
            //Use ENUMS.

            //Timings Sheet
            var dataTable = excelParser.ExcelParser("Timing", true);

            var timingLists = new List<Timing>();
         
            foreach (DataRow itemsInRow in dataTable.Rows)
            {
                timingLists.Add(new Timing
                {
                    Id = itemsInRow.ItemArray[0].ToString(),
                    Value = Convert.ToDecimal(itemsInRow.ItemArray[1]),
                    Returns = Convert.ToDecimal(itemsInRow.ItemArray[2]),
                    CurrentStatus = Convert.ToBoolean(itemsInRow.ItemArray[3].ToString().ToLower()),
                });        
            }

            return timingLists;

        }

        public static List<SectionPerformance> SectorPerformanceConverter(ExcelDataParser excelParser)
        {
            //TODO 
            //Use ENUMS.

            //SectorPerformance Sheet
            var dataTable = excelParser.ExcelParser("SectorPerformance", true);

            var sectorPerformanceLists = new List<SectionPerformance>();

            foreach (DataRow itemsInRow in dataTable.Rows)
            {
                sectorPerformanceLists.Add(new SectionPerformance
                {
                    NameEn = itemsInRow.ItemArray[0].ToString(),
                    NameFr = itemsInRow.ItemArray[1].ToString(),
                    Early = Convert.ToInt32(itemsInRow.ItemArray[2]),
                    Mid = Convert.ToInt32(itemsInRow.ItemArray[3]),
                    Late = Convert.ToInt32(itemsInRow.ItemArray[4]),
                    Recession = Convert.ToInt32(itemsInRow.ItemArray[5]),
                });
            }

            return sectorPerformanceLists;



        }




        //JSON CONVERTER.

        public static string JsonConverter(Factors factor)
        {
            string output = JsonConvert.SerializeObject(factor);

            return output;

        }










    }
}
