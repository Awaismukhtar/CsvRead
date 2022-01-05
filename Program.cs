using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

class StartUp
{
    static void Main()
    {
        Excel.Application excel = null;
        excel = new Excel.Application();
        excel.Visible = false;
        string filePath = @"C:\Users\HP\Downloads\work-Copy.xlsx";
        Excel.Workbook wkb = null;
        wkb = Open(excel, filePath);
        int value1 = 0;
        //int value2 = 24931884;

        Console.WriteLine("Enter Number to search:");
        value1 = Convert.ToInt32(Console.ReadLine());

        // string part12 = string.Concat(part1, part2);
        // List<string> lookForList = new List<string> { part1, part2, part12 };
        Excel.Range currentFind = null;
        Excel.Range searchedRange = excel.get_Range("A1", "XFD1048576");
        currentFind = searchedRange.Find(value1);
        if (currentFind != null)
        {
            Console.WriteLine("Found:");
            Console.WriteLine("Column No :" + currentFind.Column);
            Console.WriteLine("Row No :" + currentFind.Row);
        }
        else
        {
            Console.WriteLine("Not Found:");
        }
        wkb.Close(true);
        excel.Quit();
        Console.ReadLine();
    }

    public static Excel.Workbook Open(Excel.Application excelInstance,
                            string fileName, bool readOnly = false, bool editable = true,
                            bool updateLinks = true)
    {
        Excel.Workbook book = excelInstance.Workbooks.Open(
            fileName, updateLinks, readOnly,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, editable, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);
        return book;
    }
}