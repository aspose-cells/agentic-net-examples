using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineLocationDemo
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

        // Define the initial location range (cell E1) for the sparkline group
        CellArea initialLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group with the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, initialLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline at the initial location (row 0, column 4 -> E1)
        int sparkIndex = group.Sparklines.Add("A1:D1", 0, 4);
        Sparkline spark = group.Sparklines[sparkIndex];
        Console.WriteLine($"Initial position - Row: {spark.Row}, Column: {spark.Column}");

        // Desired new position: row 2 (third row), column 5 (F3)
        // Sparkline.Row and Sparkline.Column are read‑only, so we remove and re‑add the sparkline
        group.Sparklines.RemoveAt(sparkIndex);
        sparkIndex = group.Sparklines.Add("A1:D1", 2, 5);
        spark = group.Sparklines[sparkIndex];
        Console.WriteLine($"New position - Row: {spark.Row}, Column: {spark.Column}");

        // Save the workbook with the repositioned sparkline
        workbook.Save("SparklineLocationDemo.xlsx");
    }
}