using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelParserTest
{
    public class Description
    {
        public string Return { get; set; }
        public DateTime PeriodStartDate { get; set; }
        public DateTime PeriodEndDate { get; set; }
        public string RebalanceFrequency { get; set; }
        public string RankingMethod { get; set; }
        public decimal Slippage { get; set; }
        public string TransactionType { get; set; }
        public string Universe { get; set; }
        public string BenchMark { get; set; }
        public int NumberOfBuckets { get; set; }
        public decimal MinimumPrice { get; set; }
        public string Id { get; set; }
        public string NameEn { get; set; }
        public string NameFr { get; set; }
        public string DescriptionEn { get; set; }
        public string DescriptionFr { get; set; }
        public string PerformanceDescriptionEn { get; set; }
        public string PerformanceDescriptionFr { get; set; }
        public string FamilyNameEn { get; set; }
        public string FamilyNameFr { get; set; }
        public int FamilyRank { get; set; }
        public int BpsAlpha { get; set; }
        public int AdvanceIndicatorMonthMin { get; set; }
        public int AdvanceIndicatorMonthMax { get; set; }
        public int LaggingIndicatorMonthMin { get; set; }
        public int LaggingIndicatorMonthMax { get; set; }

    }

}
