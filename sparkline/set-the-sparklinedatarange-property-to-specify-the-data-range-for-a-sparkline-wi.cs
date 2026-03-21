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

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(3);
        sheet.Cells["A3"].PutValue(7);
        sheet.Cells["A4"].PutValue(2);
        sheet.Cells["A5"].PutValue(9);

        // Define the cell where the sparkline will be placed (B1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 1,
            EndColumn = 1
        };

        // Add a sparkline group with an initial vertical data range
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A5", true, location);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Access the first sparkline in the group
        Sparkline spark = group.Sparklines[0];

        // Set a new data range for the sparkline
        spark.DataRange = "A2:A5";

        // Display the updated data range
        Console.WriteLine("Updated Sparkline DataRange: " + spark.DataRange);

        // Save the workbook
        workbook.Save("SparklineDataRangeDemo.xlsx");
    }
}