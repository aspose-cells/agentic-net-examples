using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that the sparkline will represent
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(3);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(2);

        // Define the cell area where the sparkline will be placed (E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // Column E (0‑based index)
            EndColumn = 4
        };

        // Add a sparkline group of type Line with the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group using SparklineCollection.Add(string dataRange, int row, int column)
        // Use a fully‑qualified data range (including sheet name) for clarity
        string dataRange = sheet.Name + "!A1:D1";
        int sparklineIndex = group.Sparklines.Add(dataRange, 0, 4); // Row 0, Column 4 (E1)

        // Optional: customize the sparkline appearance
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Save the workbook in XLSX format
        workbook.Save("SparklineExample.xlsx", SaveFormat.Xlsx);
    }
}