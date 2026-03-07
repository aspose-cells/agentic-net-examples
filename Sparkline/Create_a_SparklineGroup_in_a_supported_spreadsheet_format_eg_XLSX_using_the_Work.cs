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

        // Populate sample data for the sparkline (vertical range A1:A5)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(3);
        sheet.Cells["A3"].PutValue(7);
        sheet.Cells["A4"].PutValue(2);
        sheet.Cells["A5"].PutValue(9);

        // Define the cell area where the sparkline will be placed (cell B1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 1,
            EndColumn = 1
        };

        // Add a sparkline group of type Line using the overload that specifies data range,
        // orientation (true = vertical), and location range.
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A5", true, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Optional: customize the appearance of the sparkline group
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Save the workbook to an XLSX file
        workbook.Save("SparklineGroupDemo.xlsx", SaveFormat.Xlsx);
    }
}