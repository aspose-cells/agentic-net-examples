using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook (XLSB will be set on save)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (row 0, columns A‑D)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell where the sparkline will be placed (E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,   // Column E (0‑based index)
            EndColumn = 4
        };

        // Add a sparkline group:
        // - Type: Line
        // - Data range: A1:D1
        // - Horizontal orientation (isVertical = false)
        // - Placement: location defined above
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Add a sparkline to the group (same data range, placed at row 0, column 4)
        int sparkIdx = group.Sparklines.Add("A1:D1", 0, 4);
        Sparkline spark = group.Sparklines[sparkIdx];

        // ---- Formatting ----
        // Set series color
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;

        // Show high and low points with custom colors
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;
        group.HighPointColor.Color = Color.Green;
        group.LowPointColor.Color = Color.Red;

        // Apply a preset style (optional)
        group.PresetStyle = SparklinePresetStyleType.Style5;

        // Save the workbook as XLSB
        workbook.Save("SparklineExample.xlsb", SaveFormat.Xlsb);
    }
}