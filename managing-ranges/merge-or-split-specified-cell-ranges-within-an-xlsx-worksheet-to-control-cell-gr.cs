using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace MergeSplitDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Header");
            cells["B1"].PutValue("Header");
            cells["C1"].PutValue("Header");
            cells["A2"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["C2"].PutValue(30);
            cells["A3"].PutValue("Detail");
            cells["B3"].PutValue("Detail");
            cells["C3"].PutValue("Detail");
            cells["D3"].PutValue("Detail");

            // -------------------------------------------------
            // Merge cells A1:C2 (rows 0-1, columns 0-2)
            // Merge(int firstRow, int firstColumn, int totalRows, int totalColumns)
            // -------------------------------------------------
            cells.Merge(0, 0, 2, 3); // merges A1:C2 into a single cell

            // -------------------------------------------------
            // Unmerge a different range B3:D3 (row 2, columns 1-3)
            // First merge it to demonstrate unmerge, then unmerge it
            // -------------------------------------------------
            cells.Merge(2, 1, 1, 3); // merge B3:D3
            cells.UnMerge(2, 1, 1, 3); // unmerge the same range

            // -------------------------------------------------
            // Use Range object to merge and then unmerge another area
            // -------------------------------------------------
            // Create a range object for E1:F2 (rows 0-1, columns 4-5)
            AsposeRange range = cells.CreateRange(0, 4, 2, 2);
            range.PutValue("Range Merge", false, false); // put a value before merging
            range.Merge(); // merge the range

            // Unmerge the same range using the Range method
            range.UnMerge();

            // -------------------------------------------------
            // Save the workbook to an XLSX file
            // -------------------------------------------------
            workbook.Save("MergeSplitResult.xlsx");
        }
    }
}