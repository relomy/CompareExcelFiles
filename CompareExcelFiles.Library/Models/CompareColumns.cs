using Aspose.Cells;

namespace CompareExcelFiles.Library.Models
{
    public class CompareColumns
    {
        public (string, string) ColumnNames { get; set; }

        public (int, int) ColumnIndices { get; set; }

        public (Workbook, Workbook) Workbooks { get; set; }
    }
}
