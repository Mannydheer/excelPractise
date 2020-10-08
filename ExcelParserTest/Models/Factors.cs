using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelParserTest.Models
{
    public class Factors
    {
        //Most likely a List of each.
        public Description Description { get; set; }
        public Performance Performance{ get; set; }
        public Combinations Combinations { get; set; }
        //4 items per list for each factor.  
        public List<Timing> Timing { get; set; } = new List<Timing>();
        public List<SectionPerformance> SectorPerformance { get; set; } = new List<SectionPerformance>();
    }

    //Combinations
    public class Combinations
    {
        public List<Contributions> Contributions { get; set; } = new List<Contributions>();
        public Correlations Correlations { get; set; }
    }

    public class Contributions
    {
        public string NameEn { get; set; }
        public string NameFr { get; set; }
        public decimal? Value { get; set; }

    }
    //Matrix.
    public class Correlations
    {
        //Better names?
        public List<string> FactorFamilyRows { get; set; }
        public List<string> FactorFamilyColumns { get; set; }
        public List<decimal> Values { get; set; }

    }
    //Dopuble check with Eric for Contributions. 
/*    public class Name
    {
        public string Momentum { get; set; }
        public int LongTermReversal { get; set; }
        public int BookToPriceRatio { get; set; }
        public int ResidualVolatility { get; set; }
        public string InvestmentQuality { get; set; }
        public decimal DividendYield { get; set; }
        public decimal Beta { get; set; }
        public decimal EarningsYield { get; set; }
        public decimal EarningsQuality { get; set; }
        public decimal Profitability { get; set; }
        public decimal Liquidity { get; set; }
        public int Size { get; set; }
    }*/

    //PERFORMANCE
    public class Performance
    {
        public string Description { get; set; }
        public RankingPerformance RankingPerformance { get; set; }
        public List<ChartDatas> ChartDatas { get; set; } = new List<ChartDatas>();
    }

    public class ChartDatas
    {
        //probably instantiate 5 times since there are 5 range.
        //Don't know if we will use ID or Name to identify.
        public string Name { get; set; }
        public List<DateTime> Periods { get; set; }
        public List<decimal> Datas { get; set; } = new List<decimal>();

    }

    public class RankingPerformance
    {
        public string Direction { get; set; }
        //Ranking Bar Graph.
        public RankingData RankingData { get; set; }
        //Own model for RankingExtraData below?
        public string Formula { get; set; }
        public decimal? Min { get; set; }
        public decimal? Median { get; set; }
        public decimal? Mean { get; set; }
        public decimal? Max { get; set; }
        public decimal? First { get; set; }
        public decimal? Last { get; set; }
        public decimal? Delta { get; set; }
        public decimal? Slope { get; set; }
        public decimal? StandardDeviation { get; set; }
    }

    public class RankingData
    {
        public List<string> Fields { get; set; } = new List<string>();

        public List<List<object>> Datas = new List<List<object>>();
    }



}
