using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetSparklineDataRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in column A (A1:A5)
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1);
        }

        // Define the cell where the sparkline will be placed (B1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 1,
            EndColumn = 1
        };

        // Add a sparkline group with a vertical data range (A1:A5)
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A5", true, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Access the first sparkline in the group
        Sparkline sparkline = group.Sparklines[0];

        // Set the DataRange property to specify the source data range (optional)
        sparkline.DataRange = "A1:A5";

        // Output the set data range (optional)
        Console.WriteLine("Sparkline DataRange set to: " + sparkline.DataRange);

        // Save the workbook to an XLSX file
        workbook.Save("SparklineDataRangeDemo.xlsx");
    }
}