using Microsoft.AspNetCore.Mvc;
using TranslateLib;
using TranslateLib.Interface;

namespace TranslateWebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class TranslateController : ControllerBase
    {

        private readonly ILogger<TranslateController> _logger;
        ITranslateExcel _translate;
        public TranslateController(ILogger<TranslateController> logger,ITranslateExcel translate)
        {
            _logger = logger;
            _translate = translate;
        }

        [HttpPost("TranslateFolder")]
        public async Task TranslateFolder(string path)
        {
            DirectoryInfo d = new DirectoryInfo(path);
            var listTasks = new List<Task>();
            var files = d.GetFiles();
            try
            {
                //foreach (var file in files)
                //{
                //    listTasks.Add(TranslateFile(file.FullName));
                //}
                //await Task.WhenAll(listTasks);
                int maxWorkerThreads = 20; // Set the maximum number of worker threads
                int maxCompletionPortThreads = 20; // Set the maximum number of IO completion port threads
                ThreadPool.SetMaxThreads(maxWorkerThreads, maxCompletionPortThreads);
                await Task.Run(() => Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = 5 },
                                    x => _translate.TranslateExcelByPathSavePath(x.FullName)
                                    ));

            }
            catch (Exception cc)
            {
                throw;
            }
        }
    }
}