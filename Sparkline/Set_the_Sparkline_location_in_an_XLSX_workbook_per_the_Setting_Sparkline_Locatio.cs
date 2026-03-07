using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in the first row (A1:E1)
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[0, i].PutValue(i + 1);
        }

        // Define the initial location range for the sparkline (cell F1)
        CellArea initialLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 5,
            EndColumn = 5
        };

        // Add a sparkline group with the data range A1:E1 and the initial location
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, initialLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline at row 0, column 5 (cell F1)
        int sparklineIndex = group.Sparklines.Add("A1:E1", 0, 5);
        Sparkline sparkline = group.Sparklines[sparklineIndex];

        // Display the current location
        Console.WriteLine($"Initial location - Row: {sparkline.Row}, Column: {sparkline.Column}");

        // To move the sparkline, remove it and add a new one at the desired location (row 2, column 5)
        group.Sparklines.RemoveAt(sparklineIndex);
        sparklineIndex = group.Sparklines.Add("A1:E1", 2, 5);
        sparkline = group.Sparklines[sparklineIndex];

        // Display the new location
        Console.WriteLine($"New location - Row: {sparkline.Row}, Column: {sparkline.Column}");

        // Save the workbook
        workbook.Save("SparklineLocationDemo.xlsx");
    }
}