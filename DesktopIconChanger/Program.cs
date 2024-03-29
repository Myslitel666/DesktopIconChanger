using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using IWshRuntimeLibrary;
using Newtonsoft.Json;

class Program
{
    static void Main(string[] args)
    {
        //Пути для папок с иконками и ярлыками
        string shortcutRootPath = @"C:\Users\user\Desktop\";
        string iconRootPath = @"D:\My Folder\Design\ico\Black-Red\";

        // Переход в директорию с JSON-файлом
        string solutionDirectory = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
        string jsonFilePath = Path.Combine(solutionDirectory, "Data", "shortcuts.json");
        // Чтение данных из JSON-файла
        string json = System.IO.File.ReadAllText(jsonFilePath);
        List<ShortcutIconPair> pairs = JsonConvert.DeserializeObject<List<ShortcutIconPair>>(json);

        WshShell shell = new WshShell();

        foreach (var pair in pairs)
        {
            string shortcutPath = shortcutRootPath + pair.Shortcut + ".lnk";
            string iconPath = iconRootPath + pair.Icon + ".ico";

            // Проверяем существование ярлыка
            if (System.IO.File.Exists(shortcutPath))
            {
                // Проверяем существование иконки
                if (System.IO.File.Exists(iconPath))
                {
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                    shortcut.IconLocation = iconPath; // Устанавливаем путь к иконке
                    shortcut.Save(); // Сохраняем изменения

                    Console.WriteLine($"Иконка ярлыка '{shortcutPath}' успешно изменена.");
                }
                else
                {
                    Console.WriteLine($"Иконка '{iconPath}' не существует, поэтому для ярлыка '{shortcutPath}' не была изменена");
                }
            }
            else
            {
                Console.WriteLine($"Ярлыка '{shortcutPath}' не существует.");
            }
        }

        Console.ReadLine();
    }

    // Класс для хранения пар "ярлык-иконка"
    class ShortcutIconPair
    {
        public string Shortcut { get; set; }
        public string Icon { get; set; }
    }
}
