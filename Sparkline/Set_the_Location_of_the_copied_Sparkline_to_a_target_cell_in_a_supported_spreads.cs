using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineCopyDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Populate sample data that the sparkline will use
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // -------------------------------------------------
            // 2. Define the location range for the original sparkline
            // -------------------------------------------------
            // The sparkline will be placed in cell E1 (row 0, column 4)
            CellArea originalLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // -------------------------------------------------
            // 3. Add a sparkline group with the data range "A1:D1"
            // -------------------------------------------------
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,          // type of sparkline
                "A1:D1",                     // data range
                false,                       // isVertical
                originalLocation);           // where the sparkline(s) will be placed

            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // -------------------------------------------------
            // 4. Add the original sparkline (automatically created by Add above)
            //    The first sparkline is at index 0
            // -------------------------------------------------
            Sparkline originalSparkline = group.Sparklines[0];

            // -------------------------------------------------
            // 5. Copy the sparkline to a new target cell (e.g., G3)
            //    Since Sparkline.Row and Sparkline.Column are read‑only,
            //    we create a new sparkline with the same DataRange at the
            //    desired location using SparklineCollection.Add.
            // -------------------------------------------------
            string dataRange = originalSparkline.DataRange; // "A1:D1"
            int targetRow = 2;    // zero‑based index for row 3 (G3)
            int targetColumn = 6; // zero‑based index for column G

            // Add a new sparkline to the same group at the target location
            int newSparklineIndex = group.Sparklines.Add(dataRange, targetRow, targetColumn);
            Sparkline copiedSparkline = group.Sparklines[newSparklineIndex];

            // -------------------------------------------------
            // 6. Output positions to verify
            // -------------------------------------------------
            Console.WriteLine($"Original sparkline - Row: {originalSparkline.Row}, Column: {originalSparkline.Column}");
            Console.WriteLine($"Copied sparkline   - Row: {copiedSparkline.Row}, Column: {copiedSparkline.Column}");

            // -------------------------------------------------
            // 7. Save the workbook in XLSX format
            // -------------------------------------------------
            workbook.Save("SparklineCopyDemo.xlsx");
        }
    }
}