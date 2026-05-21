using Aspose.Cells;
using Aspose.Cells.Charts;

class CheckSecondaryAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Series 1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Series 2");
        worksheet.Cells["C2"].PutValue(100);
        worksheet.Cells["C3"].PutValue(200);
        worksheet.Cells["C4"].PutValue(300);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Add two series and set the category axis data
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Check whether a secondary value axis already exists
        bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);

        // If the secondary axis does not exist, plot the second series on it
        if (!hasSecondaryValueAxis)
        {
            chart.NSeries[1].PlotOnSecondAxis = true;
        }

        // Ensure the secondary axis is visible (optional)
        chart.SecondValueAxis.IsVisible = true;

        // Save the workbook
        workbook.Save("CheckSecondaryAxis.xlsx");
    }
}