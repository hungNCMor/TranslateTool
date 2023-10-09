using TranslateLib.Model;

namespace TranslateLib.Interface
{
    public interface ITranslateExcel
    {
        public Task<MemoryStream> TranslateExcelByPath(string path);
        public Task TranslateExcelByPathSavePath(string path);
        public Task<MemoryStream> TranslateExcelByStream(MemoryStream stream, string fileName);
    }
}
