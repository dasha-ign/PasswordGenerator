using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.IO;

namespace PasswordGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"Список спецсимволов: {PasswordCreator.symbols}");
            //string password =  PasswordCreator.GeneratePassword(6);
            //Console.WriteLine(password.ToString());
            Console.WriteLine("Введите количество");
            int count = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Введите длину пароля");
            int length = Convert.ToInt32(Console.ReadLine());
            var passwordList = PasswordCreator.PasswordList(count, length);
            //foreach (var password in passwordList)
            //    Console.WriteLine(password);
            Console.WriteLine("Записываю список сгенерированных паролей в файл passwords.xlsx,\n располагающийся в той же папке, что и сама программа генератора");
            PasswordCreator.WriteToExcel(passwordList);
            Console.WriteLine("Запись файла завершена");
            if (File.Exists("passwords.xlsx"))
            {
                Process process = new Process();
                process.StartInfo.UseShellExecute = true;
                process.StartInfo.FileName = "passwords.xlsx";
                process.Start();
            }
                
        }

        
    }
}
