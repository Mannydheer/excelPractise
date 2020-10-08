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
        public int Slippage { get; set; }
        public string TransactionType { get; set; }
        public string Universe { get; set; }
        public string BenchMark { get; set; }
        public int NumberOfBuckets { get; set; }
        public decimal MinimumPrice { get; set; }
        public int Id { get; set; }
        public string NameFr { get; set; }
        public string NameEn { get; set; }
        public string DescriptionFr { get; set; }
        public string DescriptionEn { get; set; }
        public string FamilyNameFr { get; set; }
        public string FamilyNameEn { get; set; }
        public int FamilyRank { get; set; }
        public int BpsAlpha { get; set; }
        public int AdvanceIndicationMonthsMin { get; set; }
        public int AdvanceIndicationMonthsMax { get; set; }
       
    }

}
