using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class CopySparklineExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (A1:E1)
        for (int col = 0; col < 5; col++)
        {
            sheet.Cells[0, col].PutValue(col + 1); // values 1,2,3,4,5
        }

        // Define the location cell for the original sparkline (F1)
        CellArea originalLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 5, // column F (0‑based index)
            EndColumn = 5
        };

        // Add a sparkline group with the data range and location
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,          // sparkline type
            "A1:E1",                     // data range
            false,                       // plot by column (false = by row)
            originalLocation);           // where the sparkline will be placed

        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the first sparkline (the Add method also creates the sparkline at the location)
        int sparklineIndex = group.Sparklines.Add("A1:E1", 0, 5);
        Sparkline originalSparkline = group.Sparklines[sparklineIndex];

        // Copy the sparkline to a new location (G1) by adding another sparkline
        // with the same data range but a different column index.
        int copiedSparklineIndex = group.Sparklines.Add(
            originalSparkline.DataRange, // reuse the same data range
            0,                           // same row (row 0)
            6);                          // column G (index 6)

        // Optional: verify that the copy was created
        Sparkline copiedSparkline = group.Sparklines[copiedSparklineIndex];
        Console.WriteLine($"Original sparkline at ({originalSparkline.Row}, {originalSparkline.Column})");
        Console.WriteLine($"Copied sparkline at ({copiedSparkline.Row}, {copiedSparkline.Column})");

        // Save the workbook with the original and copied sparklines
        workbook.Save("CopySparkline.xlsx");
    }
}