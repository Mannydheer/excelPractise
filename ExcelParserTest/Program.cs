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
                
            var description = DescriptionConverter();
            var timing = TimingConverter();
            var sectorPerformance = SectorPerformanceConverter();
            var combinations = CombinationsConverter();
            var performance = PerformanceConverter();

            //Instantiate a Factor object which will be converted to JSON. 
            var factors = new Factors
            {
                Description = description,
                Timing = timing,
                SectorPerformance = sectorPerformance,
                Combinations = combinations,
                Performance = performance
            };

            var factorsJsonData = JsonConverter(factors);
         
        }

    
        public static Description DescriptionConverter()
        {
            //Pass the sheet name and if we want headerRow or not. 
            var excelReader = new ExcelDataParser();
            var dataTable = excelReader.ExcelParser("Description", false);

            var descriptionList = new List<string>();
            
            //Loop through each row.
            foreach (DataRow itemsInRow in dataTable.Rows)
            {
                //Only in the case of Period header. Since there are 3 column values.
                //In the model, two props (PeriodStateDate & PeriodEndDate) were added. 
                if (itemsInRow.ItemArray[0].ToString() == "Period" && itemsInRow.ItemArray[2] != DBNull.Value)
                {
                    descriptionList.Add(itemsInRow.ItemArray[1].ToString());
                    descriptionList.Add(itemsInRow.ItemArray[2].ToString());
                }
                else
                {
                    descriptionList.Add(itemsInRow.ItemArray[1].ToString());
                }
            }
        
             return new Description {
                    Return = descriptionList[0],
                    PeriodStartDate = DateTime.Parse(descriptionList[1]),
                    PeriodEndDate = DateTime.Parse(descriptionList[2]),
                    RebalanceFrequency = descriptionList[3],
                    RankingMethod = descriptionList[4],
                    Slippage = Convert.ToDecimal((descriptionList[5])),
                    TransactionType = descriptionList[6],
                    Universe = descriptionList[7],
                    BenchMark = descriptionList[8],
                    NumberOfBuckets = Convert.ToInt32(descriptionList[9]),
                    MinimumPrice = Convert.ToDecimal(descriptionList[10]),
                    Id = descriptionList[11],
                    NameEn = descriptionList[12],
                    NameFr = descriptionList[13],
                    DescriptionEn = descriptionList[14],
                    DescriptionFr = descriptionList[15],
                    PerformanceDescriptionEn = descriptionList[16],
                    PerformanceDescriptionFr = descriptionList[17],
                    FamilyNameEn = descriptionList[18],
                    FamilyNameFr = descriptionList[19],
                    FamilyRank = Convert.ToInt32(descriptionList[20]),
                    BpsAlpha = Convert.ToInt32(descriptionList[21]),
                    AdvanceIndicatorMonthMin = Convert.ToInt32(descriptionList[22]),
                    AdvanceIndicatorMonthMax = Convert.ToInt32(descriptionList[23]),
                    //Changed value to 0 need to check with if they will always provide values. 
                    LaggingIndicatorMonthMin = Convert.ToInt32(descriptionList[24]),
                    LaggingIndicatorMonthMax = Convert.ToInt32(descriptionList[25]),
             };
        }

        public static Performance PerformanceConverter()
        {
            //TODO need to add description prop for perofrmance.
            var performance = new Performance();
            
            performance.RankingPerformance = RankingPerformanceConverter();
            performance.ChartDatas = ChartDatasConverter();

            return performance;
        }

        //Ranking_ExtraData & Ranking Sheet.
        public static RankingPerformance RankingPerformanceConverter()
        {
            var excelReader = new ExcelDataParser();

            //Ranking_ExtraData SHeet.
            var dataTableForRankingExtraDataSheet = excelReader.ExcelParser("Ranking_ExtraData", true);

            RankingPerformance rankingPerformance = null;

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

            //Ranking Sheet
            //Now need to isntantiate ranking dat ana dfill the fields. 
            var dataTableForRankingSheet = excelReader.ExcelParser("Ranking", false);
       
            List<object> rankingDataListsFirstRow = dataTableForRankingSheet.Rows[0].ItemArray.ToList();

            List<object> rankingDataListsSecondRow = dataTableForRankingSheet.Rows[1].ItemArray.ToList();

            //Start at 1 to skip first column of each row.
            for (int i = 1; i < rankingDataListsFirstRow.Count(); i++)
            {
                //Construct array of obj of key-valye.
                rankingPerformance.RankingData.Add(new RankingData
                {
                    Field = rankingDataListsFirstRow[i].ToString(),
                    Value = Convert.ToDecimal(rankingDataListsSecondRow[i])
                });
            }

            return rankingPerformance;
        }

        //TODO
        //Performance Sheet. 
        public static ChartDatas ChartDatasConverter()
        {
            var excelReader = new ExcelDataParser();

            //TODO
            //Use ENUMS.
            //Timings Sheet
            var dataTable = excelReader.ExcelParser("Performance", false);

            var chartData = new ChartDatas();

            var isFirstRow = true;

            foreach (DataRow itemsInRow in dataTable.Rows)
            {

                if (isFirstRow)
                {
                    isFirstRow = false;

                    for (int i = 0; i < itemsInRow.ItemArray.Count(); i++)
                    {
                        //first field date it empty...
                        chartData.Fields.Add(itemsInRow.ItemArray[i].ToString());
                    }
                }
                else
                {
                    var dataList = new List<string>();

                    for (int i = 0; i < itemsInRow.ItemArray.Count(); i++)
                    {
                        dataList.Add(itemsInRow.ItemArray[i].ToString());
                    }
                    chartData.Datas.Add(dataList);
                }
               
                
            }


            return chartData;

        }

        //Combinations Sheet. 

        public static Combinations CombinationsConverter()
        {
            var excelReader = new ExcelDataParser();
            var dataTable = excelReader.ExcelParser("Contributions", true);

            var combinations = new Combinations();

            //Contributions Sheet
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
            var dataTableCorrelations = excelReader.ExcelParser("Correlation", false);


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

        public static List<Timing> TimingConverter()
        {
            var excelReader = new ExcelDataParser();

            //TODO
            //Use ENUMS.
            //Timings Sheet
            var dataTable = excelReader.ExcelParser("Timing", true);

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

        public static List<SectionPerformance> SectorPerformanceConverter()
        {
            var excelReader = new ExcelDataParser();

            //TODO 
            //Use ENUMS.
            //SectorPerformance Sheet

            var dataTable = excelReader.ExcelParser("SectorPerformance", true);

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
            //TODO
            //TODO [10:50 AM] Benoit Parenteau
            //JsonNamingPolicy.CamelCase


            return JsonConvert.SerializeObject(factor);
        }










    }
}
