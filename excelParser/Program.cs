using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Reflection;

namespace excelParser
{


    class Program
    {
        static void Main(string[] args)
        {
            //Item a = new Item();
            var a = typeof(Item);
            var b = a.GetFields();

            Dictionary<string, int?> map = new Dictionary<string, int?>();

            foreach (var fieldInfo in b)
            {
                map.Add(fieldInfo.Name, null);
            }

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Neerd\source\repos\excel\excelParser\excelParser\ex.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            List<Array> listLines = new List<Array>();

            Console.WriteLine("///////////");
            Console.WriteLine(xlRange.Rows.Count);
            Console.WriteLine(xlRange.Columns.Count);
            Console.WriteLine("///////////");



            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;


            for (int i = 1; i <= colCount; i++)
            {
                string c = xlRange.Cells[1, i].Value2.ToString();
                if (map.ContainsKey(c))
                {
                    map[c] = i;
                }
            }




            List<Item> items = new List<Item>();

            for (int i = 2; i <= rowCount; i++)
            {
                Item item = new Item();

                foreach (var mappedElement in map)
                {
                    FieldInfo fieldInfo = typeof(Item).GetField(mappedElement.Key);

                    fieldInfo.SetValue(item, xlRange.Cells[i, mappedElement.Value].Value2);

                }

                items.Add(item);

                //string[] myArr = new string[colCount];
                //string line = "";
                //string tempLine = "";
                //for (int j = 1; j <= colCount; j++)
                //{
                //    //new line
                //    if (j == 1)
                //        Console.Write("\n");

                //    //write the value to the console
                //    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                //    {
                //        // Each elem in console
                //        // Console.Write(xlRange.Cells[i, j].Value2.ToString() + " ");
                //        tempLine = xlRange.Cells[i, j].Value2.ToString();
                //        myArr[j-1] = tempLine;
                //        Console.WriteLine(" ADD TO myArr: {0}", tempLine);
                //    }
                //    line = line + tempLine + " ";

                //}
                //listLines.Add(myArr);
                //Console.WriteLine(line);
            }

            Console.WriteLine("--------------LIST LINES----------------");
            foreach (var item in items)
            {
                
                Console.WriteLine($"{item.Prop1} | {item.Prop2} | {item.Prop3}");
            }


            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //close excel
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.FinalReleaseComObject(xlWorkbook);
            //Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            //Marshal.ReleaseComObject(xlApp);

            Process.GetProcessesByName("EXCEL").ToList().ForEach(x =>
            {
                x.Kill();
            });

        }
    }
}
