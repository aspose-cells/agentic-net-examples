using Aspose.Cells;
using Aspose.Cells.Charts;

class DisableDataLabelAutoFit
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Set Y values
        chart.NSeries.CategoryData = "A2:A4";      // Set X categories

        // Enable data labels for the first series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Disable auto‑fit (auto‑resize) for each data label and set a fixed width
        foreach (ChartPoint point in series.Points)
        {
            point.DataLabels.IsResizeShapeToFitText = false; // Turn off auto‑fit
            point.DataLabels.WidthPixel = 60;                // Fixed width in pixels
        }

        // Save the workbook
        workbook.Save("DisableDataLabelAutoFit.xlsx");
    }
}