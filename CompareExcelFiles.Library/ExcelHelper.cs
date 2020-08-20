using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Cells;
using CompareExcelFiles.Library.Models;

namespace CompareExcelFiles.Library
{
    public static class ExcelHelper
    {
        public static readonly License License = new License();

        static ExcelHelper()
        {
            License.SetLicense("Aspose.Total.lic");
        }

        public static Workbook OpenFile(string path)
        {
            // open workbook
            Workbook workbook = new Workbook(path);

            if (workbook.Worksheets == null || workbook.Worksheets.Count == 0)
            {
                return null;
            }

            return workbook;
        }

        public static Dictionary<string, int> GetHeaders(Workbook workbook)
        {
            var headers = new Dictionary<string, int>();

            // get first worksheet
            var worksheet = workbook.Worksheets[0];

            // iterate through the first row through the max column
            for (int i = 0; i <= worksheet.Cells.MaxColumn; i++)
            {
                var cell = worksheet.Cells.GetCell(0, i);

                // skip empty cells
                if (cell == null || cell.Value == null)
                {
                    continue;
                }

                headers.Add(cell.Value.ToString(), cell.Column);
            }

            return headers;
        }

        public static void Compare(CompareColumns compareColumns)
        {
            System.Console.WriteLine($"Comparing {compareColumns.ColumnNames.Item1} and {compareColumns.ColumnNames.Item2}");
            var worksheet1 = compareColumns.Workbooks.Item1.Worksheets[0];
            var worksheet2 = compareColumns.Workbooks.Item2.Worksheets[0];

            // for each row for worksheet1 / column 1, search for it in worksheet2
            for (int i = 1; i <= worksheet1.Cells.MaxDataRow; i++)
            {
                var cellToFind = worksheet1.Cells.GetCell(i, compareColumns.ColumnIndices.Item1);

                if (cellToFind == null || cellToFind.Value == null)
                {
                    System.Console.WriteLine($"[{i}, {compareColumns.ColumnIndices.Item1}] cellToFind is null / cellToFind.Value is null");
                    continue;
                }

                // look for the cell's value in worksheet2
                for (int j = 1; j < worksheet2.Cells.MaxDataRow; j++)
                {
                    var cellToSearch = worksheet2.Cells.GetCell(j, compareColumns.ColumnIndices.Item2);

                    if (cellToSearch == null || cellToSearch.Value == null)
                    {
                        System.Console.WriteLine($"[{j}, {compareColumns.ColumnIndices.Item2}] cellToFind is null / cellToFind.Value is null");
                        continue;
                    }

                    if (cellToFind.Value.ToString().Equals(cellToSearch.Value.ToString()))
                    {
                        System.Console.WriteLine($"found a MATCH! [{i}, {compareColumns.ColumnIndices.Item1}] '{cellToFind.Value}' == " +
                            $"'[{j}, {compareColumns.ColumnIndices.Item2}] '{cellToSearch.Value}'");
                    }
                }
            }
        }
    }
}