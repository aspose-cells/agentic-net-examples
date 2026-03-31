using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineLocationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location range where sparklines will be placed (initially column E, row 1)
            CellArea locationRange = new CellArea
            {
                StartRow = 0,   // Row index 0 (first row)
                EndRow = 0,
                StartColumn = 4, // Column index 4 (E)
                EndColumn = 4
            };

            // Add a sparkline group with the data range A1:D1
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, locationRange);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add a sparkline at the specified location (row 0, column 4 -> cell E1)
            // Using the rule: SparklineCollection.Add(string dataRange, int row, int column)
            int sparklineIdx = group.Sparklines.Add("A1:D1", 0, 4);
            Sparkline sparkline = group.Sparklines[sparklineIdx];

            Console.WriteLine($"Initial Sparkline position -> Row: {sparkline.Row}, Column: {sparkline.Column}");

            // To change the sparkline's position, remove the existing sparkline
            group.Sparklines.RemoveAt(sparklineIdx);

            // Add the sparkline again at a new location (row 2, column 4 -> cell E3)
            sparklineIdx = group.Sparklines.Add("A1:D1", 2, 4);
            sparkline = group.Sparklines[sparklineIdx];

            Console.WriteLine($"New Sparkline position -> Row: {sparkline.Row}, Column: {sparkline.Column}");

            // Save the workbook to an XLSX file
            workbook.Save("SparklineLocationDemo.xlsx");
        }
    }
}