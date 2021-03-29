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

            var a = ExcelParser.DesFile<Item>(@"D:\1REPOS\excelParser\excelParser\ex.xls");

            Console.WriteLine("--------------LIST LINES----------------");
            foreach (var item in a)
            {
                Console.WriteLine($"{item.NAME} | {item.PRICE}| {item.BRAND}");
            }

        }
    }
}
