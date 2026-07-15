using Aspose.Cells;
using Aspose.Cells.Charts;

class ChartPositionExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Add some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a chart; initial position will be adjusted later
        int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 3, 15, 3);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Move the chart so its upper‑left corner is at row 15, column 3
        // Bottom row and right column are set to the same values; the actual size will be defined by WidthPt
        chart.Move(15, 3, 15, 3);

        // Set the chart width to 400 points
        chart.ChartObject.WidthPt = 400;

        // Save the workbook
        workbook.Save("ChartPositioned.xlsx");
    }
}