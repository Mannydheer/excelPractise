using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelParserTest.Models
{
    public class SectionPerformance
    {
        public string NameEn { get; set; }
        public string NameFr { get; set; }
        public int Early { get; set; }
        public int Mid { get; set; }
        public int Late { get; set; }
        public int Recession { get; set; }

    }

    public enum PerformancePattern
    {
        NoClearPattern,
        ConsistentlyUnderperform,
        Underperform,
        Overperform,
        ConsistentlyOverperform

    }
}
/*public int Communication { get; set; }
public int ConsumerDictionary { get; set; }
public int ConsumerStaples { get; set; }
public int Energy { get; set; }
public int Financials { get; set; }
public int HealthCare { get; set; }
public int Industrials { get; set; }
public int InfoTechnology { get; set; }
public int Materials { get; set; }
public int RealEstate { get; set; }
public int Utilies { get; set; }*/