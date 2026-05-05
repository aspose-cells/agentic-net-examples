using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate first row with sample data for the original sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(2);

            // Define the location where the original sparkline will be placed (cell E1)
            CellArea originalLocation = CellArea.CreateCellArea("E1", "E1");

            // Add a sparkline group with the original data range and location
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // The Add method already created a sparkline at E1; retrieve it if needed
            Sparkline originalSparkline = group.Sparklines[0];

            // ----- Copy the sparkline -----
            // Add new data for the copied sparkline (second row)
            sheet.Cells["A2"].PutValue(7);
            sheet.Cells["B2"].PutValue(4);
            sheet.Cells["C2"].PutValue(6);
            sheet.Cells["D2"].PutValue(3);

            // Define the location for the copied sparkline (cell F1)
            // Note: we add a new sparkline to the same group with the same initial data range
            int copiedIdx = group.Sparklines.Add("A1:D1", 0, 5); // row 0, column 5 => F1
            Sparkline copiedSparkline = group.Sparklines[copiedIdx];

            // Set the data range of the copied sparkline to the second row data
            copiedSparkline.DataRange = "A2:D2";

            // Optional: display the data ranges in console for verification
            Console.WriteLine("Original Sparkline DataRange: " + originalSparkline.DataRange);
            Console.WriteLine("Copied Sparkline DataRange: " + copiedSparkline.DataRange);

            // Save the workbook
            workbook.Save("SparklineCopyWithNewDataRange.xlsx");
        }
    }
}