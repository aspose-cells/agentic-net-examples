using Aspose.Cells;
using Aspose.Cells.Charts;

class HideScatterGridlines
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the scatter plot (X and Y values)
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue(2);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue(4);
        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["B4"].PutValue(6);

        // Add a scatter chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the Y values for the series and bind the X values
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries[0].XValues = "A2:A4";

        // Hide the horizontal (category) axis gridlines for a cleaner look
        chart.CategoryAxis.MajorGridLines.IsVisible = false;
        chart.CategoryAxis.MinorGridLines.IsVisible = false;

        // Save the workbook with the modified chart
        workbook.Save("ScatterNoGridlines.xlsx");
    }
}