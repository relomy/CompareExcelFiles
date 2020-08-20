using System;
using System.Collections.Generic;
using CompareExcelFiles.Library;
using CompareExcelFiles.Library.Models;

namespace CompareExcelFiles.ConsoleUI
{
    public class Program
    {
        static void Main(string[] args)
        {
            const string filePath1 = @"C:\Users\alewando\Documents\exceltest\test1.xlsx";
            const string filePath2 = @"C:\Users\alewando\Documents\exceltest\test2.xlsx";

            // open excel files
            Console.WriteLine($"Opening {filePath1}");
            var workbook1 = ExcelHelper.OpenFile(filePath1);
            var headers1 = ExcelHelper.GetHeaders(workbook1);
            PrintHeaders(headers1);

            Console.WriteLine($"Opening {filePath2}");
            var workbook2 = ExcelHelper.OpenFile(filePath2);
            var headers2 = ExcelHelper.GetHeaders(workbook2);
            PrintHeaders(headers2);

            // ask user about which headers to compare

            // headers picked, look at selected columns for matches
            List<CompareColumns> compareColumns = new List<CompareColumns>
            {
                new CompareColumns()
                {
                    ColumnNames = ("last name", "last name"),
                    ColumnIndices = (headers1["last name"], headers2["last name"]),
                },
                new CompareColumns()
                {
                    ColumnNames = ("first name", "first name"),
                    ColumnIndices = (headers1["first name"], headers2["first name"]),
                },
            };


            // display results
        }

        public static void PrintHeaders(Dictionary<string, int> headers)
        {
            foreach (var header in headers)
            {
                Console.WriteLine($"Key: {header.Key} Value: {header.Value}");
            }
        }
    }
}
