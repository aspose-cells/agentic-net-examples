using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class SparklineDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for two rows (A1:E2)
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[0, i].PutValue(i + 1);          // Row 0: 1,2,3,4,5
            sheet.Cells[1, i].PutValue(5 - i);          // Row 1: 5,4,3,2,1
        }

        // Define the location where sparklines will be placed (column F, rows 0‑1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 1,
            StartColumn = 5,
            EndColumn = 5
        };

        // Add a sparkline group of type Line for the data range A1:E2
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E2", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Add a sparkline for each row of the data range
        group.Sparklines.Add(sheet.Name + "!A1:E1", 0, 5); // Sparkline for row 0
        group.Sparklines.Add(sheet.Name + "!A2:E2", 1, 5); // Sparkline for row 1

        // Apply a preset style and enable high/low point markers
        group.PresetStyle = SparklinePresetStyleType.Style5;
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Set custom colors for high and low points
        CellsColor highColor = workbook.CreateCellsColor();
        highColor.Color = Color.Green;
        group.HighPointColor = highColor;

        CellsColor lowColor = workbook.CreateCellsColor();
        lowColor.Color = Color.Red;
        group.LowPointColor = lowColor;

        // Save the workbook with the sparklines
        workbook.Save("SparklinesDemo.xlsx");
    }
}