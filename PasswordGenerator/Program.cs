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
            int count = 1;
            int length = 12;
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Введите количество");
                count = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine("Введите длину пароля");
                length = Convert.ToInt32(Console.ReadLine());
            }
            else
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if ((args[i] == "-h" || args[i] == "-help")
                        ||
                        (args[i] == "\\h" || args[i] == "\\help"))
                    {
                        Console.WriteLine("Ключи запуска");
                        Console.WriteLine("-l,-lenght,\\l,\\lenght - длина генерируемого пароля");
                        Console.WriteLine("-c,-count,\\c,\\count - количество паролей, которое будет сгенерировано");

                    }
                    if (((args[i] == "-c" || args[i] == "-count") 
                        ||
                        (args[i] == "\\c" || args[i] == "\\count") )
                        && (i != args.Length - 1))
                    {
                        try
                        {
                            count = Convert.ToInt32(args[i + 1]);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Не удаестся распознать нужное количество паролей с помощью ключей запуска\nБудет сгенерирован 1 пароль");
                        }
                    }
                    if (((args[i] == "-l" || args[i] == "-lenght") 
                        ||
                        (args[i] == "\\l" || args[i] == "\\lenght"))
                        && (i != args.Length - 1))
                    {
                        try
                        {
                            count = Convert.ToInt32(args[i + 1]);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Не удаестся распознать длину пароля с помощью ключей запуска\nБудет сгенерирован пароль длиной 12 символов");
                        }
                    }

                }
            }
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
