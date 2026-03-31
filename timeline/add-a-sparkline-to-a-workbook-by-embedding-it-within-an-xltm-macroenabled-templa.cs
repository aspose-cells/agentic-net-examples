using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the macro‑enabled template (XLTM)
        Workbook workbook = new Workbook("Template.xltm");
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define where the sparkline will be placed (cell E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group of type Line using the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,
            sheet.Name + "!A1:D1",
            false,
            location);

        // Retrieve the created group (the sparkline is already added by the Add method)
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Optionally, add another sparkline to the same group
        // group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Save the workbook, preserving the macro‑enabled template format
        workbook.Save("OutputTemplate.xltm");
    }
}