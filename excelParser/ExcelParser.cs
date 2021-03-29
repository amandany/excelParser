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
    public static class ExcelParser
    {
        public static List<T> DesFile<T>(string path) where T : new()
        {
            var a = typeof(T);
            var b = a.GetFields();

            Dictionary<string, int?> map = new Dictionary<string, int?>();

            foreach (var fieldInfo in b)
            {
                map.Add(fieldInfo.Name, null);
            }

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            List<Array> listLines = new List<Array>();

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


            List<T> items = new List<T>();

            for (int i = 2; i <= rowCount; i++)
            {
                T item = new T();
                if (xlRange.Cells[i, 1]?.Value2 == null)
                {
                    continue;
                }
                foreach (var mappedElement in map)
                {
                    FieldInfo fieldInfo = typeof(T).GetField(mappedElement.Key);

                    fieldInfo.SetValue(item, xlRange.Cells[i, mappedElement.Value].Value2);

                }

                items.Add(item);
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

            return items;

        }
    }
}
