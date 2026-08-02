using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the sparkline (row 4)
        sheet.Cells["A4"].PutValue(5);
        sheet.Cells["B4"].PutValue(2);
        sheet.Cells["C4"].PutValue(1);
        sheet.Cells["D4"].PutValue(3);

        // Define the location cell I4 (column index 8, row index 3)
        CellArea location = new CellArea
        {
            StartRow = 3,
            EndRow = 3,
            StartColumn = 8,
            EndColumn = 8
        };

        // Add a line sparkline group with the data range A4:D4 placed at I4
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A4:D4", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Enable highlighting of the first and last points
        group.ShowFirstPoint = true;
        group.ShowLastPoint = true;

        // Optional: set colors for the first and last points
        CellsColor firstPointColor = workbook.CreateCellsColor();
        firstPointColor.Color = Color.Purple;
        group.FirstPointColor = firstPointColor;

        CellsColor lastPointColor = workbook.CreateCellsColor();
        lastPointColor.Color = Color.Yellow;
        group.LastPointColor = lastPointColor;

        // Save the workbook
        workbook.Save("SparklineFirstLastPoint.xlsx");
    }
}