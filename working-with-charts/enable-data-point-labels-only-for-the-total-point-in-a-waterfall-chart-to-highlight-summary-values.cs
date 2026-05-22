using Aspose.Cells;
using Aspose.Cells.Charts;

class WaterfallTotalLabelDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a waterfall chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Start");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Increase");
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["A4"].PutValue("Decrease");
        sheet.Cells["B4"].PutValue(-20);
        sheet.Cells["A5"].PutValue("Total");
        sheet.Cells["B5"].PutValue(0); // placeholder; Excel will compute the total

        // Add a Waterfall chart
        int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Mark the total point (index 4, zero‑based) as a subtotal (i.e., a total) 
        chart.NSeries[0].LayoutProperties.Subtotals = new int[] { 4 };

        // Ensure all points are generated before accessing them
        chart.Calculate(new ChartCalculateOptions() { UpdateAllPoints = true });

        // Enable data label only for the total point
        Series series = chart.NSeries[0];
        foreach (int subtotalIndex in series.LayoutProperties.Subtotals)
        {
            ChartPoint totalPoint = series.Points[subtotalIndex];
            totalPoint.DataLabels.ShowValue = true;                     // show the numeric value
            totalPoint.DataLabels.Position = LabelPositionType.OutsideEnd; // position the label
            totalPoint.DataLabels.IsAutoText = false;                  // allow custom text
            totalPoint.DataLabels.Text = "Total";                       // custom label text
        }

        // Save the workbook
        workbook.Save("WaterfallTotalLabel.xlsx");
    }
}