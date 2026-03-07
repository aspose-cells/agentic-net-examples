using Aspose.Cells;
using Aspose.Cells.Charts;

class InsertSparklineXlsb
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell area where the sparkline will be placed (E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group of type Line, linking it to the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // (Optional) Add an explicit sparkline; the Add method above already creates one
        // group.Sparklines.Add("A1:D1", 0, 4);

        // Save the workbook as an XLSB file using XlsbSaveOptions
        XlsbSaveOptions saveOptions = new XlsbSaveOptions();
        saveOptions.ExportAllColumnIndexes = true; // ensure all column indexes are exported

        workbook.Save("SparklineExample.xlsb", saveOptions);
    }
}