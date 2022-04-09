using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace PasswordGenerator
{
    class PasswordCreator
    {
        const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        const string digits = "0123456789";
        public const string symbols = "!@#()_+-="; // "!@#$%^&*()-_=+<,>.";
    //    string allSymbols = chars + chars.ToLower() + digits + symbols;

        public static List<string> PasswordList(int count, int passwordLength)
        {
            var result = new List<string>();
            do
            {
                string password = GeneratePassword(passwordLength);
                if (!result.Contains(password))
                    result.Add(password);
            }
            while (result.Count < count);
            return result;
        }

        public static string GeneratePassword(int length = 13, List<string> existingPasswords = null)
        {
            string allSymbols = chars + chars.ToLower() + digits + symbols;
            //4 Capital letters, 4 small letters, 4 digits & 1 special symbol
            char[] password = new char[length];
            var random = new Random();
            var randomNumbers = Enumerable.Range(0, length).OrderBy(x => random.Next()).Take(length).ToList();
            // we need to ensure first password symbol is going to be character
            if (randomNumbers.ToArray()[0] == length-1) 
                randomNumbers = randomNumbers.ToArray().Reverse().ToList();
            int specSymbolPosition = random.Next(1, length - 1);
            foreach (var number in randomNumbers)
            {
                _ = randomNumbers.IndexOf(number) switch
                {
                    int i when (i >= 0 && i <= 7)  =>
                        password[number] = randomNumbers.IndexOf(number) % 2 == 0 ?
                                            chars[random.Next(0, 25)] :
                                            chars.ToLower()[random.Next(0, chars.Length)],
                    int i when i > 7 && i < 11 =>
                        password[number] = digits[random.Next(0, digits.Length)],
                    int i when i == 11 =>
                    password[number] = symbols[random.Next(0, symbols.Length)],
                    _ => password[number] =  allSymbols[random.Next(0, allSymbols.Length)]
                };
            }
            return new string(password);
        }

        public static void WriteToExcel(List<string> passwordList)
        {
            using var workBook = new XLWorkbook();
            {

                var workSheet = workBook.Worksheets.Add("Пароли");
                workSheet.Cell("A1").Value = "Список паролей";
                int i = 2;
                foreach (string password in passwordList)
                {
                    workSheet.Cell(i, 1).Value = password;
                    i++;
                }
                workSheet.Column(1).AdjustToContents();
                workBook.SaveAs("passwords.xlsx");
            }
        }
    }
}
