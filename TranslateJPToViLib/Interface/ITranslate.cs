using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TranslateLib.Interface
{
    public interface ITranslate
    {
        public string TranslateText(string input, string src = "ja", string des = "vi");
    }
}
