using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace МодЭкзамен
{ 
    /// <summary>
    /// Программа критического пути
    /// </summary>
    class Program
    {
        /// <summary>
        /// Главная программа запуска
        /// </summary>
        /// <param name="args"> Параматры для программ </param>
        static void Main(string[] args)
        {
        iz2: string listname = string.Empty;
            Console.WriteLine("Какой лист хотите выбрать?");
            switch (Console.ReadLine())
            {
                case "1":
                    Console.WriteLine("Выбран 1 лист. Подождите минуту.");
                    listname = "Лист1";
                    break;
                default:
                    Console.WriteLine("Такого листа нет.");
                    break;
            }
            string filename = @"D:\Программирование\C#\МодЭкзамен\КритПуть.XLSX";
            int column = 0;
            double[,] table = Excel.GetArray(filename, listname, out column);

            int n;
            int[] i = new int[20];
            int[] j = new int[20];
            int[] dij = new int[20];
            int[] s1 = new int[20];
            int[] s2 = new int[20];
            int[] f1 = new int[20];
            int[] f2 = new int[20];
            int[] tf = new int[20];
            int[] ff = new int[20];

            n = table.GetLength(0);
            Log.WriteLog($"Общее количество работ: {n}");
            for (int k = 0; k < n; k++)
            {
                i[k] = (int)table[k, 0];
                j[k] = (int)table[k, 1];
                dij[k] = (int)table[k, 2];
            }

            Critical_Path(n, i, j, dij, s1, s2, f1, f2, tf, ff);

            Console.WriteLine("Самый ранний срок начала: ");
            for (int k = 0; k < n; k++)
            {
                Console.Write("{0} \n", s1[k]);
            }
            Console.WriteLine("Самый поздний срок начала: ");
            for (int k = 0; k < n; k++)
            {
                Console.Write("{0} \n", s2[k]);
            }
            Console.WriteLine("Самый ранний срок завершения: ");
            for (int k = 0; k < n; k++)
            {
                Console.Write("{0} \n", f1[k]);
            }
            Console.WriteLine("Самый поздний срок завершения: ");
            for (int k = 0; k < n; k++)
            {
                Console.Write("{0} \n", f2[k]);
            }
            Console.WriteLine("Свободный резерв времени: ");
            for (int k = 0; k < n; k++)
            {
                Console.Write("{0} \n", ff[k]);
            }
            Console.WriteLine("Полный резерв времени: ");
            for (int k = 0; k < n; k++)
            {
                Console.Write("{0} \n", tf[k]);
            }
            string str = "1";
            Console.WriteLine("Критический путь:");
            Console.Write(str);
            for (int k = 0; k < n; k++)
            {
                if (tf[k] == 0)
                    Console.Write(" {0}",j[k]);
            }

            Excel.ExportToExcel(filename, listname, n, s1, s2, f1, f2, tf, ff, j);

        iz1: Console.WriteLine("\nХотите продолжить работу? (y/n)");
            var exit = Console.ReadLine();
            if (exit.ToLower() != "y" && exit.ToLower() != "n")
            {
                Console.WriteLine("Ответ неверный, повторите попытку.");
                goto iz1;
            }
            if (exit.ToLower() == "y")
            {
                Console.Clear();
                goto iz2;
            }
        }

        /// <summary>
        /// Вычисление критического пути
        /// </summary>
        /// <param name="n"> общее количество работ по проекту</param>
        /// <param name="i"> вектор пара представляющую К работу</param>
        /// <param name="j"> вектор пара которая понимается как стрелка</param>
        /// <param name="dij"> продолжительность К операции</param>
        /// <param name="s1"> самый ранний срок начала К операции</param>
        /// <param name="s2">самый поздний срок началаК операции</param>
        /// <param name="f1">самый ранний срок завершения К операции</param>
        /// <param name="f2">самый поздний срок завершения К операции</param>
        /// <param name="tf">Полный резерв времени К операции</param>
        /// <param name="ff">Свободный резерв времени К операции</param>
        public static void Critical_Path(int n, int[] i, int[] j, int[] dij, int[] s1, int[] s2, int[] f1, int[] f2, int[] tf, int[] ff)
        {
            int k, index, max =int.MinValue, min =int.MaxValue;
            int[] ti = new int[20];
            int[] te = new int[20];

            index = 0;

            for (k = 0; k < n; k++)
            {
                if (i[k] == index + 1) index = i[k];
                ti[k] = 0;
                te[k] = 9999;
            }

            for (k = 0; k < n; k++)
            {
                max = ti[i[k]] + dij[k];
                if (ti[j[k]] < max) ti[j[k]] = max;
            }
            Log.WriteLog($"Максимум: {max}");

            te[j[n - 1]] = ti[j[n - 1]];

            for (k = n - 1; k >= 0; k--)
            {
                min = te[j[k]] - dij[k];
                if (te[i[k]] > min) te[i[k]] = min;
            }
            Log.WriteLog($"Минимум: {min}");
            for (k = 0; k < n; k++)
            {
                s1[k] = ti[i[k]];
                f1[k] = s1[k] + dij[k];
                f2[k] = te[j[k]];
                s2[k] = f2[k] - dij[k];
                tf[k] = f2[k] - f1[k];
                ff[k] = ti[j[k]] - f1[k];
            }
        }
    }
   /// <summary>
   /// Логирование событий при работе программы
   /// </summary>
    public class Log
    {
        static string filename = @"D:\Программирование\C#\МодЭкзамен\log.txt";
        /// <summary>
        /// Добавление одного сообщения в файл
        /// </summary>
        /// <param name="message"> сообщение для добавления</param>
        public static void WriteLog(string message)
        {
            using (StreamWriter sw = File.AppendText(filename))
            {
                sw.WriteLine(message);
            }
        }
    }
    /// <summary>
    /// Класс для импорта и экспорта в эксель
    /// </summary>
    public class Excel
    {
        /// <summary>
        /// Импорт данных из эксекль
        /// </summary>
        /// <param name="filename"> путь к эксель файлу</param>
        /// <param name="listname">название листа в эксель</param>
        /// <param name="column">количество колонок импортируемых</param>
        /// <returns>возвращение двухмерного массива данных</returns>
        public static double[,] GetArray(string filename, string listname, out int column)
        {
            Application xlApp = new Application();
            Workbook xlWB;
            Worksheet xlSht;
            column = 0;
            xlWB = xlApp.Workbooks.Open(filename);
            xlSht = xlWB.Worksheets[listname];
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "B"].End[XlDirection.xlUp].Row;
            object[,] arrData = xlSht.Range["A2:Z" + iLastRow].Value;
            for (var i = 1; i <= iLastRow - 2; i++)
            {
                for (var j = 1; j <= arrData.GetLength(1); j++)
                {
                    if (arrData[i, j] == null) continue;
                    column++;
                }
                break;
            }
            double[,] table = new double[iLastRow - 2, column];
            for (var i = 1; i <= iLastRow - 2; i++)
            {
                for (var j = 1; j <= column; j++)
                {
                    table[i - 1, j - 1] = Convert.ToDouble(arrData[i, j]);
                    Console.WriteLine("\t " + table[i - 1, j - 1]);
                }
                Console.WriteLine("\n");
            }
            xlWB.Close(false);
            xlApp.Quit();
            return table;
        }
        /// <summary>
        /// Экспорт данных в Эксель
        /// </summary>
        /// <param name="filename"> путь к эксель файлу</param>
        /// <param name="listname">название листа в эксель</param>
        /// <param name="n"> общее количество работ по проекту</param>
        /// <param name="j"> вектор пара которая понимается как стрелка</param>
        /// <param name="s1"> самый ранний срок начала К операции</param>
        /// <param name="s2">самый поздний срок началаК операции</param>
        /// <param name="f1">самый ранний срок завершения К операции</param>
        /// <param name="f2">самый поздний срок завершения К операции</param>
        /// <param name="tf">Полный резерв времени К операции</param>
        /// <param name="ff">Свободный резерв времени К операции</param>
       
        public static void ExportToExcel(string filename, string listname, int n, int[] s1, int[] s2, int[] f1, int[] f2, int[] tf, int[] ff, int[] j)
        {
            Application excelApp = new Application();
            excelApp.Visible = true;
            Workbook xlWB;
            Worksheet xlSht;
            xlWB = excelApp.Workbooks.Open(filename, XlUpdateLinks.xlUpdateLinksNever, false);
            xlSht = xlWB.Worksheets[listname];
            var str = string.Empty;
            xlSht.Cells[12, "A"] = "Самый ранний срок начала: ";
            for (int k = 0; k < n; k++)
            {
                str+= " " + s1[k];
            }
            xlSht.Cells[13, "A"] = str;
            str = string.Empty;

            xlSht.Cells[15, "A"] = "Самый поздний срок начала: ";
            for (int k = 0; k < n; k++)
            {
                str += " " + s2[k];
            }

            xlSht.Cells[16, "A"] = str;
            str = string.Empty;

            xlSht.Cells[18, "A"] = "Самый ранний срок завершения: ";
            for (int k = 0; k < n; k++)
            {
                str += " " + f1[k];
            }

            xlSht.Cells[19, "A"] = str;
            str = string.Empty;

            xlSht.Cells[21, "A"] = "Самый поздний срок завершения: ";
            for (int k = 0; k < n; k++)
            {
                str += " " + f2[k];
            }

            xlSht.Cells[22, "A"] = str;
            str = string.Empty;

            xlSht.Cells[24, "A"] = "Свободный резерв времени: ";
            for (int k = 0; k < n; k++)
            {
                str += " " + ff[k];
            }

            xlSht.Cells[25, "A"] = str;
            str = string.Empty;

            xlSht.Cells[27, "A"] = "Полный резерв времени: ";
            for (int k = 0; k < n; k++)
            {
                str += " " + tf[k];
            }

            xlSht.Cells[28, "A"] = str;
            str = string.Empty;

            str = "1 ";
            xlSht.Cells[29, "A"] = "Критический путь: ";
            for (int k = 0; k < n; k++)
            {
                if (tf[k]==0) str += j[k] + " ";
            }
            xlSht.Cells[30, "A"] = str;

            excelApp.DisplayAlerts = false;
            xlSht.SaveAs(string.Format(@"D:\Программирование\C#\МодЭкзамен\КритПуть.XLSX", Environment.CurrentDirectory));

            excelApp.Quit();
        }
    }
}
