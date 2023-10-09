using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TranslateLib.Model
{
    public class DataResult
    {
        public string Value { get; set; }
        public bool IsFormula { get; set; } = false;
        public int row { get; set; }
        public int column { get; set; }
        public int Sheet { get; set; }
    }
}
