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

        //public static string GetColumnHeading(string docName, string worksheetName, string cellName)

        //{
        //    // Open the document as read-only.
        //    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))

        //    {
        //        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);

        //        if (sheets.Count() == 0)

        //        {
        //            // The specified worksheet does not exist.

        //            return null;
        //        }

        //        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);

        //        // Get the column name for the specified cell.

        //        string columnName = GetColumnName(cellName);

        //        // Get the cells in the specified column and order them by row.

        //        IEnumerable<Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => string.Compare(GetColumnName(c.CellReference.Value), columnName, true) == 0)

        //            .OrderBy(r => GetRowIndex(r.CellReference));

        //        if (cells.Count() == 0)

        //        {
        //            // The specified column does not exist.

        //            return null;
        //        }

        //        // Get the first cell in the column.

        //        Cell headCell = cells.First();

        //        // If the content of the first cell is stored as a shared string, get the text of the first cell

        //        // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.

        //        if (headCell.DataType != null && headCell.DataType.Value == CellValues.SharedString)

        //        {
        //            SharedStringTablePart shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();

        //            SharedStringItem[] items = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();

        //            return items[int.Parse(headCell.CellValue.Text)].InnerText;
        //        }
        //        else

        //        {
        //            return headCell.CellValue.Text;
        //        }
        //    }
        //}

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

        public static Compare((int, int) columnIndices)
        {

        }

        // Given a cell name, parses the specified cell to get the column name.
        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);
            return match.Value;
        }

        // Given a cell name, parses the specified cell to get the row index.

        private static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);
            return uint.Parse(match.Value);
        }
    }
}