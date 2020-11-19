using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Security;
using System.Threading;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            string WorkPath = Environment.CurrentDirectory;
            //EXCEL

            DirectoryInfo dInfo = new DirectoryInfo(WorkPath);
            int x = dInfo.GetFiles("*.xls*").Length;
            string[] files = Directory.GetFiles(WorkPath,"*.xls*");
            Console.WriteLine("Всего файлов: "+x.ToString());
            Console.WriteLine("С какой строки идут заголовки?Если их нет нажмите Enter");
            int startstr = 1;
            bool contin = false;
            while (contin != true)
            {
                try
                {
                    Console.WriteLine("Введите номер строки с заголовками:");
                    string num = Console.ReadLine();
                    if (num == "")
                    {
                        break;
                    }
                    else
                    { 
                        startstr = Convert.ToInt32(num);
                        contin = true;
                    }
                }
                catch
                {
                    Console.WriteLine("Неверное значение!");
                    Console.WriteLine("Продолжить или ввести заново?(Y/N)");
                    string confirm = Console.ReadLine();
                    if(confirm == "Y" || confirm == "y" || confirm == "н" || confirm == "Н")
                    {
                        contin = false;
                    }
                    else
                    {
                        startstr = 1;
                        contin = true;                        
                    }
                }
            }
            string outname = "";
            Console.WriteLine("Введите имя выходного файла:");
            while (outname == "")
            {
                outname = Console.ReadLine().Replace(' ','_') ;
                if (outname == "")
                {
                    Console.WriteLine("Вы ввели пустое имя файла");
                }
            }

            Excel.Application mainEx = new Excel.Application();
            Excel.Workbook mWorkBook = mainEx.Workbooks.Add(1);
            Excel.Worksheet mWorkSheet = (Excel.Worksheet)mWorkBook.Sheets[1];

            mainEx.Visible = false;
            mainEx.ScreenUpdating = false;
            mainEx.DisplayAlerts = false;
            int nfile = 1;
            foreach (string str in files)
            {
                try
                {
                    Excel.Application secEx = new Excel.Application();
                    secEx.Visible = false;
                    secEx.ScreenUpdating = false;
                    secEx.DisplayAlerts = false;
                    Excel.Workbook sWorkBook = secEx.Workbooks.Open(str, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Excel.Worksheet sWorkSheet = (Excel.Worksheet)sWorkBook.Sheets[1];

                    string urange = sWorkSheet.UsedRange.Address;
                    int contstr = startstr + 1;
                    if (startstr == 1)
                    {
                        sWorkSheet.UsedRange.Copy();
                    }
                    else
                    {
                        if (nfile == 1)
                        {
                            sWorkSheet.Rows[startstr].Copy();
                            mWorkSheet.Paste();
                            mWorkSheet.Range["A2", "A2"].Select();
                        }
                        string[] rangecells = urange.Split(':');
                        sWorkSheet.Range["A" + (contstr).ToString(), rangecells[1]].Copy();
                    }    
                    mWorkSheet.Paste();
                    int r = mWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                    //r = r.Replace('O', 'A');

                    mWorkSheet.Range["A" + (r-1).ToString(), "A" + (r - 1).ToString()].Select();
                    Console.WriteLine("Ипортировано: " + str +"(" + nfile.ToString() + " из " + x + ")");
                    sWorkBook.Close();
                    secEx.Quit();
                }
                catch(Exception ex)
                {
                    Console.WriteLine("Ошибка с " + str + "(" + nfile.ToString() + " из " + x + ")");
                    Console.WriteLine(ex.Message.ToString());
                }
                nfile += 1;
            }

            mWorkBook.SaveAs(WorkPath+ @"\" + outname + ".xlsx");
            mWorkBook.Close();
            mainEx.Quit();
            Console.WriteLine("Готово, нажмите Enter чтобы закрыть приложение.");
            Console.ReadLine();



            //mWorkBook.SaveAs(WorkPath+ @"\main.xlsx");
            //mWorkBook.Close();
            //mainEx.Quit();
        }
    }
}
