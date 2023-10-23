using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TranslateLib.Interface
{
    public interface ITranslateFile
    {
        public Task<MemoryStream> TranslateFileByPath(string path);
        public Task TranslateFileByPathSavePath(string path);
        public Task<MemoryStream> TranslateFileByStream(MemoryStream stream, string fileName);
    }
}
