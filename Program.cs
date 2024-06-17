using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace TaskCSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Поиск макс. длины массива
            int minStringLength = 10;
            int stringLength = minStringLength;
            long currentArraySize = stringLength * 2; //* 2 б
            while (currentArraySize <= 2000000000) //2 Гб
            {
                stringLength += 1;
                currentArraySize += stringLength * 2;
            }
            int maxArrayLength = stringLength - minStringLength;

            // Ввод размера массива
            uint N = 0;
            Console.Write("Введите количество элементов массива (N): ");
            while (!UInt32.TryParse(Console.ReadLine(), out N) || (N > maxArrayLength)) 
            {
                Console.Clear();
                Console.Write($"N должно быть целым числом от 0 до {maxArrayLength}.\nВведите количество элементов массива (N): ");
            };


            //Указание пути сохранения
            Console.Write(@"Введите путь к месту сохранения файла (Например, C:\Users\User\Desktop): ");
            string path = Console.ReadLine();
            while (!Directory.Exists(path) || path.EndsWith(@"\"))
            {
                Console.Write(@"Указан неправильный путь. Введите путь к месту сохранения файла (Например, C:\Users\User\Desktop): ");
                path = Console.ReadLine();
            }

            //Определение названия файла
            string fileName =  path + @"\SortedData-" + DateTime.Now.ToString("yy-MMMM-dd HH-mm-ss") + ".xlsx";
            Console.WriteLine(fileName);


            //Генерация строк
            string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzабвгдеёжзийклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ0123456789";
            string[] strings = new string[N];
            Random rand = new Random();
            for (int n = 0; n < N; n++) //цикл по массиву
            {
                do
                {
                    strings[n] = "";
                    for (int i = 0; i < n+10; i++) //цикл по строке тек. элемента массива
                        strings[n] += alphabet[rand.Next(0, alphabet.Length)];
                }
                while (!strings[n].Any(s => char.IsLetter(s)) || !strings[n].Any(s => char.IsDigit(s)));
                Console.WriteLine($"{n+1}) {strings[n]}");
            };

            //Сортировка по возрастанию
            string[] sortedByAsc = new string[N];
            strings.CopyTo(sortedByAsc, 0);
            Array.Sort(sortedByAsc);
            //Сортировка по убыванию
            string[] sortedByDesc = sortedByAsc.Reverse().ToArray();


            //Создание Excel-файла
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            //Заполнение заголовков
            workSheet.Cells[1, "A"] = "№";
            workSheet.Cells[1, "B"] = "Начальные сгенерированные данные";
            workSheet.Cells[1, "C"] = "Отсортированные по возрастанию";
            workSheet.Cells[1, "D"] = "Отсортированные по убыванию";
            //Заполнение столбцов
            for (int i = 0; i < strings.Length; i++)
            {
                workSheet.Cells[i+2, "A"] = i + 1;
                workSheet.Cells[i+2, "B"] = strings[i];
                workSheet.Cells[i+2, "C"] = sortedByAsc[i];
                workSheet.Cells[i+2, "D"] = sortedByDesc[i];
            }
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
            workSheet.Columns[4].AutoFit();

            //Сохранение в файл
            excelApp.ActiveWorkbook.SaveAs(fileName);
        }
    }
}
