using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class SparklineHighLowMarkersDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(1);

        // Define the cell area where the sparkline will be placed (E1)
        CellArea location = new CellArea
        {
            StartColumn = 4, // Column E (0‑based index)
            EndColumn = 4,
            StartRow = 0,    // Row 1
            EndRow = 0
        };

        // Add a sparkline group of type Line with the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Enable markers for both high and low points
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Optional: set colors for the high and low point markers
        CellsColor highColor = workbook.CreateCellsColor();
        highColor.Color = Color.Green;
        group.HighPointColor = highColor;

        CellsColor lowColor = workbook.CreateCellsColor();
        lowColor.Color = Color.Red;
        group.LowPointColor = lowColor;

        // Save the workbook to a file
        workbook.Save("SparklineHighLowMarkers.xlsx");
    }
}