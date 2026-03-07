using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineLocationDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that the sparkline will represent
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the area where the sparkline will be placed (single cell)
            // Here we use column 4 (E) and row 0 (first row) as the initial location range.
            CellArea locationRange = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group of type Line, linking it to the data range A1:D1.
            // The group is created with the location range defined above.
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, locationRange);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Specify the exact location of a sparkline using SparklineCollection.Add(row, column).
            // This places the sparkline in cell E1 (row 0, column 4).
            int sparklineIndex = group.Sparklines.Add("A1:D1", 0, 4);
            Sparkline sparkline = group.Sparklines[sparklineIndex];

            // Output the location to verify (Row and Column are read‑only properties)
            Console.WriteLine($"Sparkline placed at row {sparkline.Row}, column {sparkline.Column}");

            // Save the workbook in XLSX format
            workbook.Save("SparklineLocationDemo.xlsx");
        }
    }
}