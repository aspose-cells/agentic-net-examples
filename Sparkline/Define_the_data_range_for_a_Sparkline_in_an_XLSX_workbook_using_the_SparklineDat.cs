using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineDataRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(3);
            sheet.Cells["A3"].PutValue(7);
            sheet.Cells["A4"].PutValue(2);
            sheet.Cells["A5"].PutValue(9);

            // Define where the sparkline will be placed (cells B1:B5)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 1,
                EndColumn = 1
            };

            // Add a sparkline group with an initial data range
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A5", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Access the first sparkline in the group
            Sparkline sparkline = group.Sparklines[0];

            // Set a new data range for the sparkline using the DataRange property
            sparkline.DataRange = "A1:A5";

            // Optional: display the data range in console
            Console.WriteLine("Sparkline DataRange set to: " + sparkline.DataRange);

            // Save the workbook
            workbook.Save("SparklineDataRangeDemo.xlsx");
        }
    }
}