using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class AddSparklineToTemplate
{
    static void Main()
    {
        // Load the XLTX template
        Workbook workbook = new Workbook("Template.xltx");

        // Get the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data that the sparkline will represent
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell where the sparkline will be placed (E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a Line sparkline group using the data range A1:D1,
        // not vertical (plot by column), and place it in the defined location.
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Optional: customize appearance of the sparkline group
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        group.HighPointColor = workbook.CreateCellsColor();
        group.HighPointColor.Color = Color.Green;

        group.LowPointColor = workbook.CreateCellsColor();
        group.LowPointColor.Color = Color.Red;

        // Save the workbook with the added sparkline
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}