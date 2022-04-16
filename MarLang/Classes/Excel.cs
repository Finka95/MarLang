using System;
using System.Text;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using Telegram.Bot.Types;

namespace MarLang
{
    public class Excel
    {
        public static Dictionary<long,int> ListUsers {get;private set;}
        public static List<List<string>>[] arrayUsersQuizWords {get;set;}
        public static void SetQuizWords(long chatId)
        {
            int index;
            ListUsers.TryGetValue(chatId, out index);
            arrayUsersQuizWords[index] = GetAllWords(chatId);
        }
        public static bool DeleteWord(long chatId ,string word)
        {
            int dfdf = 12;
            int mainIndex;
            ListUsers.TryGetValue(chatId, out mainIndex);
            using (var package = new ExcelPackage(new FileInfo(@"YOUR PATH")))
            {
                var sheet = package.Workbook.Worksheets[mainIndex];
                for (int i = 4; !string.IsNullOrEmpty(sheet.Cells[$"A{i}"].Text); i++)
                {
                    if(sheet.Cells[$"A{i}"].Text.ToLower() == word.ToLower())
                    {
                        sheet.DeleteRow(i);
                        package.Save();
                        return true;
                    }
                }
            }
            return false;
        }
        private static bool CheckArray(int[,] array, int val_1, int val_2)
        {
            for (int i = 0; i < array.GetLength(0); i++)
            {
                if(array[i,0] == val_1 && array[i,1] == val_2)
                    return true;
            }
            return false;
        }
        public static QuizClass GetQuizCard(long chatId)
        {
            int mainIndex;
            ListUsers.TryGetValue(chatId, out mainIndex);
            List<List<string>> QuizWords = arrayUsersQuizWords[mainIndex];

            Random rnd = new Random();
            int index = rnd.Next(0,QuizWords.Count);
            string word = QuizWords[index][0];
            string descriprion = QuizWords[index][rnd.Next(1,QuizWords[index].Count)];
            descriprion = descriprion.Length > 100 ? string.Format(descriprion[..96] + "...") : descriprion;
            QuizWords.RemoveAt(index);
            int [,] indices = new int [3,2];
            string[] descriptions = new string[4];
            index = rnd.Next(0,4); 
            descriptions[index] = descriprion;
            int[,] indeces = new int[4,2];
            int index_1 = default;
            int index_2 = default;
            for (int i = 0; i < 4; i++)
            {
                if(i == index)
                    continue;
                do
                {
                    index_1 = rnd.Next(0,QuizWords.Count);
                    index_2 = rnd.Next(1,QuizWords[index_1].Count);
                } while (CheckArray(indeces, index_1, index_2));
                indeces[i,0] = index_1;
                indeces[i,1] = index_2;
                descriptions[i] = QuizWords[index_1][index_2].Length > 100 ? string.Format(QuizWords[index_1][index_2][..96] + "...") : QuizWords[index_1][index_2];
            }
            
            if(QuizWords.Count < 4)
                QuizWords = null;

            arrayUsersQuizWords[mainIndex] = QuizWords;
            return new QuizClass(word, index, descriptions);
        }
        public static string GetLastTenWords(long chatId)
        {
            int mainIndex;
            ListUsers.TryGetValue(chatId, out mainIndex);
            StringBuilder result = new StringBuilder();
            using (var package = new ExcelPackage(new FileInfo(@"YOUR PATH")))
            {
                var sheet = package.Workbook.Worksheets[mainIndex];
                for (int i = 0; i < 10; i++)
                {
                    result.Append(sheet.Cells[$"A{i + 4}"].Text + "\n");
                    char colomn = 'B';
                    for (int y = 0; y < 5; y++)
                    {
                        if(string.IsNullOrEmpty(sheet.Cells[$"{colomn}{i + 4}"].Text))
                            break;
                        result.Append(sheet.Cells[$"{colomn++}{i + 4}"].Text + "\n");
                    }
                    result.Append("\n");
                }
            }
            return result.ToString();
        }
        public static void AddNewWord(long chatId, string word, List<string> descriprion)
        {
            if(CheckWordInBook(chatId, word))
                return;

            int mainIndex;
            ListUsers.TryGetValue(chatId, out mainIndex);

            using (var package = new ExcelPackage(new FileInfo(@"YOUR PATH")))
            {
                var sheet = package.Workbook.Worksheets[mainIndex];
                sheet.InsertRow(4,1);
                sheet.Cells["A4"].Value = word;
                char colomn = 'B';
                foreach (string item in descriprion)
                {
                    sheet.Cells[$"{colomn++}4"].Value = item;
                }
                package.Save();
            }
        }
        public static bool CheckWordInBook(long chatId, string word)
        {
            word = word.ToLower();
            List<List<string>> result = GetAllWords(chatId);
            foreach (var item in result)
            {
                if(item[0].ToLower() == word)
                    return true;
            }
            return false;
        }
        public static List<List<string>> GetAllWords(long chatId)
        {
            int mainIndex;
            ListUsers.TryGetValue(chatId, out mainIndex);
            var result = new List<List<string>>();
            using (var package = new ExcelPackage(new FileInfo(@"/Users/vlad/Projects/MarLang/MarLang/WordList.xlsx")))
            {
                var sheet = package.Workbook.Worksheets[mainIndex];
                for (int i = 0; !string.IsNullOrEmpty(sheet.Cells[$"A{i + 4}"].Text); i++)
                {
                    result.Add(new List<string>());
                    result[i].Add(sheet.Cells[$"A{i + 4}"].Text);
                    char colomn = 'B';
                    for (int y = i + 4; !string.IsNullOrEmpty(sheet.Cells[$"{(char)(colomn)}{y}"].Text); colomn++)
                    {
                        result[i].Add(sheet.Cells[$"{colomn}{y}"].Text);
                    }
                }
            }
            return result;
        }
        public static void GetListUsers()
        {
            Dictionary<long,int> result = new Dictionary<long, int>();
            List<IsDeletedClass> listDeleted = new List<IsDeletedClass>();
            using (var package = new ExcelPackage(new FileInfo(@"YOUR PATH")))
            {
                var sheet = package.Workbook.Worksheets[0];
                for (int i = 1; !string.IsNullOrEmpty(sheet.Cells[$"A{i}"].Text); i++)
                {
                    result.Add(long.Parse(sheet.Cells[$"A{i}"].Text),int.Parse(sheet.Cells[$"B{i}"].Text));
                    listDeleted.Add(new IsDeletedClass(long.Parse(sheet.Cells[$"A{i}"].Text)));
                }
            }
            arrayUsersQuizWords = new List<List<string>>[result.Count + 2];
            ListUsers = result;
            Handlers.isDeletedList = listDeleted;
        }
        public static int AddNewUser(long chatId)
        {
            ListUsers.Add(chatId, ListUsers.Count + 2);
            var newArray = new List<List<string>>[arrayUsersQuizWords.Length + 1];
            for (int i = 0; i < arrayUsersQuizWords.Length; i++)
                newArray[i] = arrayUsersQuizWords[i];
            arrayUsersQuizWords = newArray;
            using (var package = new ExcelPackage(new FileInfo(@"YOUR PATH")))
            {
                var sheet = package.Workbook.Worksheets[0];
                for (int i = 1; true; i++)
                {
                    if(string.IsNullOrEmpty(sheet.Cells[$"A{i}"].Text))
                    {
                        sheet.Cells[$"A{i}"].Value = chatId;
                        sheet.Cells[$"B{i}"].Value = $"{ListUsers.Count + 2}";
                        break;
                    }
                }
                package.Workbook.Worksheets.Copy("1",$"{ListUsers.Count + 2}");
                package.Save();
            }
            return 0;
        }
        public static string CreateExcelFile(long chatId)
        {
            string path = @"YOUR NEW EXCEL FILE PATH";
            int mainIndex;
            ListUsers.TryGetValue(chatId, out mainIndex);
            using (var package = new ExcelPackage(new FileInfo(@"YOUR PATH")))
            {
                using (var package1 = new ExcelPackage())
                {
                    FileInfo excelFile = new FileInfo(path);
                    package1.Workbook.Worksheets.Add("0", package.Workbook.Worksheets[mainIndex]);
                    package1.SaveAs(excelFile);
                }
            }
            return path;
        }
    }
}