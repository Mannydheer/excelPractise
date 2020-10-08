using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelParserTest.Models
{
    public class Timing
    {
        public TimingIds Id { get; set; }
        public int Value { get; set; }
        public decimal? Returns { get; set; }
        public bool CurrentStatus { get; set; }
    }

    public enum TimingIds
    {
        Early,
        Mid,
        Late,
        Recession        
    }
}
