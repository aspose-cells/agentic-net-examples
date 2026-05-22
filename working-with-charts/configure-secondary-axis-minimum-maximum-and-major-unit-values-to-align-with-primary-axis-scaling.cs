using Aspose.Cells;
using Aspose.Cells.Charts;

class ConfigureSecondaryAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for two series
        cells["A1"].PutValue("Category");
        cells["A2"].PutValue("A");
        cells["A3"].PutValue("B");
        cells["A4"].PutValue("C");

        cells["B1"].PutValue("Series1");
        cells["B2"].PutValue(100);
        cells["B3"].PutValue(200);
        cells["B4"].PutValue(300);

        cells["C1"].PutValue("Series2");
        cells["C2"].PutValue(1000);
        cells["C3"].PutValue(2000);
        cells["C4"].PutValue(3000);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series to the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Configure primary (value) axis scaling
        Axis primaryAxis = chart.ValueAxis;
        primaryAxis.IsAutomaticMinValue = false;
        primaryAxis.IsAutomaticMaxValue = false;
        primaryAxis.IsAutomaticMajorUnit = false;
        primaryAxis.MinValue = 0;          // Minimum
        primaryAxis.MaxValue = 3500;       // Maximum
        primaryAxis.MajorUnit = 500;       // Major unit

        // Align secondary axis scaling with the primary axis
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.IsAutomaticMinValue = false;
        secondaryAxis.IsAutomaticMaxValue = false;
        secondaryAxis.IsAutomaticMajorUnit = false;
        secondaryAxis.MinValue = primaryAxis.MinValue;
        secondaryAxis.MaxValue = primaryAxis.MaxValue;
        secondaryAxis.MajorUnit = primaryAxis.MajorUnit;

        // Save the workbook
        workbook.Save("AlignedSecondaryAxis.xlsx");
    }
}