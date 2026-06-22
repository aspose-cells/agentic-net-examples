using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineInCell
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the location where the sparkline will be placed (cell P5)
            // Column P = index 15, Row 5 = index 4 (zero‑based)
            CellArea location = new CellArea
            {
                StartRow = 4,
                EndRow = 4,
                StartColumn = 15,
                EndColumn = 15
            };

            // Define the data range for the sparkline (including sheet name)
            string dataRange = $"{sheet.Name}!B2:B10";

            // Add a sparkline group of type Line.
            // isVertical = false because the data is in a single column.
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, dataRange, false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // (Optional) Customize the sparkline appearance here, e.g.:
            // group.ShowHighPoint = true;
            // group.ShowLowPoint = true;

            // Save the workbook with the sparkline embedded in cell P5
            workbook.Save("SparklineInP5.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}