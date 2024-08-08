using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TJADSZY.ai
{
    internal class utilities_ai
    {
        public static List<string> ExtractUrls(string jsonString)
        {
            // 正则表达式匹配URL
            string pattern = @"https?://[^\s""]+";
            Regex regex = new Regex(pattern);

            MatchCollection matches = regex.Matches(jsonString);
            List<string> urls = new List<string>();

            foreach (Match match in matches)
            {
                urls.Add(match.Value);
            }

            return urls;
        }
    }
    
}
