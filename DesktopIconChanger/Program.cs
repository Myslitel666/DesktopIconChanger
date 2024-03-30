using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using IWshRuntimeLibrary;

class Program
{
    static void Main(string[] args)
    {
        // Путь к папке с ярлыками на общем рабочем столе
        string commonDesktopPath = @"C:\Users\Public\Desktop\";

        // Путь к папке с ярлыками на персонализированном (пользовательском) рабочем столе
        string userDesktopPath = @"C:\Users\user\Desktop\";

        // Путь к папке с иконками
        string iconRootPath = @"D:\My Folder\Design\ico\Black-Red\";

        // Получаем список файлов иконок в указанной папке
        string[] iconFiles = Directory.GetFiles(iconRootPath, "*.ico");

        // Извлекаем имена файлов без расширений и приводим их к нижнему регистру
        List<string> iconNames = iconFiles.Select(Path.GetFileNameWithoutExtension)
                                           .Select(name => name.ToLower())
                                           .ToList();

        WshShell shell = new WshShell();

        // Обновляем ярлыки на общем рабочем столе
        UpdateShortcuts(shell, commonDesktopPath, iconRootPath, iconNames);

        // Обновляем ярлыки на персонализированном рабочем столе
        UpdateShortcuts(shell, userDesktopPath, iconRootPath, iconNames);

        Console.ReadLine();
    }

    static void UpdateShortcuts(WshShell shell, string shortcutPath, string iconRootPath, List<string> iconNames)
    {
        foreach (var shortcutFilePath in Directory.GetFiles(shortcutPath, "*.lnk"))
        {
            // Получаем только имя файла без расширения и пути
            string fileName = Path.GetFileNameWithoutExtension(shortcutFilePath);

            // Приводим имя файла к нижнему регистру
            fileName = fileName.ToLower();

            // Инициализируем переменную для хранения самой длинной подстроки из списка имен иконок
            string longestIconName = "";

            // Перебираем все имена иконок из списка
            foreach (var iconName in iconNames)
            {
                // Если имя файла содержит подстроку из списка имен иконок и длина этой подстроки больше предыдущей найденной
                if (fileName.Contains(iconName) && iconName.Length > longestIconName.Length)
                {
                    // Обновляем самую длинную подстроку
                    longestIconName = iconName;
                }
            }

            // Если найдена подходящая иконка, создаем путь к ней и заменяем иконку ярлыка
            if (!string.IsNullOrEmpty(longestIconName))
            {
                // Создаем путь к соответствующей иконке
                string iconPath = Path.Combine(iconRootPath, $"{longestIconName}.ico");

                try
                {
                    // Заменяем иконку ярлыка
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutFilePath);
                    shortcut.IconLocation = iconPath; // Устанавливаем путь к иконке
                    shortcut.Save(); // Сохраняем изменения

                    Console.WriteLine($"Иконка ярлыка '{shortcutFilePath}' успешно изменена.");
                }
                catch
                {
                    Console.WriteLine($"Ошибка доступа к директории '{shortcutPath}'. Запустите приложение от имени администратора");
                }
            }
            else
            {
                // Выводим сообщение о том, что подходящая иконка не найдена
                Console.WriteLine($"Для ярлыка '{shortcutFilePath}' подходящая иконка не найдена.");
            }
        }
    }
}
