using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class CreateColumnSparkline
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the location where the sparkline will be placed (cell C5)
            // Row and column indices are zero‑based (C5 -> row 4, column 2)
            CellArea location = new CellArea
            {
                StartRow = 4,
                EndRow = 4,
                StartColumn = 2,
                EndColumn = 2
            };

            // Define the data range for the sparkline (D5:D15 on the same sheet)
            string dataRange = $"{sheet.Name}!D5:D15";

            // Add a sparkline group of type Column.
            // isVertical = false (horizontal orientation for the data series)
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Column, dataRange, false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // (Optional) Access the created sparkline if further customization is needed
            // Sparkline sparkline = group.Sparklines[0];

            // Save the workbook with the sparkline
            workbook.Save("ColumnSparkline.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}