namespace CompareExcelFiles.Library.Models
{
    public class CompareColumns
    {
        public (string, string) ColumnNames { get; set; }

        public (int, int) ColumnIndices { get; set; }
    }
}
