using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class SparklineGroupExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (row 0, columns A-D)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(3);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(2);

        // Define the location where the sparkline will be placed (cell E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // Column E (0‑based index)
            EndColumn = 4
        };

        // Add a sparkline group of type Line, using the data range A1:D1,
        // not vertical (plot by column), and place it in the defined location
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,
            "A1:D1",
            false,
            location);

        // Retrieve the created SparklineGroup
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Optionally add a sparkline explicitly (the Add method already created one,
        // but this demonstrates the Sparklines collection usage)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Customize appearance (example: set series color and show high/low points)
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;

        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Save the workbook as an XLSX file
        workbook.Save("SparklineGroupExample.xlsx", SaveFormat.Xlsx);
    }
}