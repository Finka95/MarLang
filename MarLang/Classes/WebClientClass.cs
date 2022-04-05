using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace MarLang
{
    public static class WebClientClass
    {
        public static List<string> GetDescriptions(string word)
        {
            word = word.ToLower();
            List<string> result = new List<string>();
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < word.Length; i++)
            {
                if (char.IsLetter(word[i]))
                    sb.Append(word[i]);
                else
                    sb.Append("-");
            }
            string sbString = $"https://www.oxfordlearnersdictionaries.com/definition/english/{sb.ToString()}_1?q={sb.Replace('-','+').ToString()}";
            string response;
            try
            {
                response = new WebClient().DownloadString(sbString);
            }
            catch (System.Exception)
            {
                return result;
            }
            var search = Regex.Matches(response, $"(class=\\\"def\\\" htag=\\\"span\\\">|\"span\\\" hclass=\\\"def\\\">)([a-zA-Z ,.!?;:]+)</span");
            for (int i = 0; i < search.Count && i < 5; i++)
            {
                result.Add(search[i].Groups[2].ToString());
            }
            return result;
        }

        public static bool ReturnValue(this List<IsDeletedClass> isDeletedClass, long chatId)
        {
            for (int i = 0; i < isDeletedClass.Count; i++)
            {
                if (isDeletedClass[i].chatId == chatId)
                    return isDeletedClass[i].isDeleted;
            }
            return false;
        }

        public static void SetValue(this List<IsDeletedClass> isDeletedClass, long chatId, bool isDeleted)
        {
            for (int i = 0; i < isDeletedClass.Count; i++)
            {
                if (isDeletedClass[i].chatId == chatId)
                {
                    isDeletedClass[i].isDeleted = isDeleted;
                    return;
                }
            }
        }
    }
}