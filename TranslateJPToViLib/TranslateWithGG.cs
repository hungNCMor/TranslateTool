using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using TranslateLib.Interface;

namespace TranslateJPToViLib
{
    public class TranslateWithGG : ITranslate
    {
        ILogger<TranslateWithGG> _logger;
        public TranslateWithGG(ILogger<TranslateWithGG> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// default is the ja to vi
        /// </summary>
        /// <param name="input"></param>
        /// <param name="src"></param>
        /// <param name="des"></param>
        /// <returns></returns>
        public string TranslateText(string input, string src = "ja", string des = "vi")
        {
            var t2 = DateTime.Now;
            string url = String.Format
            ("https://translate.googleapis.com/translate_a/single?client=gtx&tl={0}&sl={1}&dt=t&q={2}",
             des, src, Uri.EscapeUriString(input));
            HttpClient httpClient = new HttpClient();
            try
            {
                string responseBody = httpClient.GetStringAsync(url).Result;

                // Parse the response to get the translated text
                string translatedText = ParseTranslationResponse(responseBody);
                _logger.LogInformation("TranslateText time " + (DateTime.Now - t2));
                return translatedText;
            }
            catch (Exception xx)
            {

                throw;
            }

        }
        private string ParseTranslationResponse(string response)
        {
            dynamic data = JsonConvert.DeserializeObject(response);
            try
            {
                string extractedString = data[0][0][0].ToString();
                return extractedString;
            }
            catch (Exception c)
            {

                throw;
            }


        }
    }
}