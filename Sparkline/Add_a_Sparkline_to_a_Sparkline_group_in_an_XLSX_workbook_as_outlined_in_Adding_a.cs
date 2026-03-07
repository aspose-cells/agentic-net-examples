using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AddSparklineToGroup
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
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
        // plotting by row (isVertical = false), and placing the sparkline at E1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group.
        // Data range can be qualified with the sheet name; location is row 0, column 4 (E1)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Optional: customize the sparkline appearance
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = System.Drawing.Color.Orange;
        group.SeriesColor = seriesColor;
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Save the workbook as an XLSX file
        workbook.Save("SparklineAdded.xlsx");
    }
}