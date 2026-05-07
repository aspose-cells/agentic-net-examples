using Aspose.Cells;
using Aspose.Cells.Charts;

class ActivateSparkline3D
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that will be used for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(3);
        sheet.Cells["A3"].PutValue(7);
        sheet.Cells["A4"].PutValue(2);
        sheet.Cells["A5"].PutValue(9);

        // Define where the sparkline will be placed (B1:B5)
        CellArea location = CellArea.CreateCellArea("B1", "B5");

        // Add a sparkline group of type Line, using the data range A1:A5
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A5", false, location);
        SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

        // Note: 3‑D rendering for sparklines is not supported in this version of Aspose.Cells.
        // The sparkline group is created and will be displayed in the workbook.

        // Save the workbook to an XLSX file
        workbook.Save("Sparkline3D.xlsx");
    }
}