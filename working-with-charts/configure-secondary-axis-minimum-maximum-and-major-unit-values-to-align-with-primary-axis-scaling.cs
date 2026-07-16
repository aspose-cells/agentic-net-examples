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

        // Populate sample data
        cells["A1"].PutValue("Category");
        cells["A2"].PutValue("A");
        cells["A3"].PutValue("B");
        cells["A4"].PutValue("C");

        cells["B1"].PutValue("Primary");
        cells["B2"].PutValue(100);
        cells["B3"].PutValue(200);
        cells["B4"].PutValue(300);

        cells["C1"].PutValue("Secondary");
        cells["C2"].PutValue(1000);
        cells["C3"].PutValue(2000);
        cells["C4"].PutValue(3000);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series: first uses primary axis, second uses secondary axis
        chart.NSeries.Add("B2:B4", true); // primary series
        chart.NSeries.Add("C2:C4", true); // secondary series
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Configure primary value axis scaling
        Axis primaryAxis = chart.ValueAxis;
        primaryAxis.IsAutomaticMinValue = false;
        primaryAxis.IsAutomaticMaxValue = false;
        primaryAxis.IsAutomaticMajorUnit = false;
        primaryAxis.MinValue = 0;
        primaryAxis.MaxValue = 400;
        primaryAxis.MajorUnit = 100;

        // Configure secondary value axis to match primary axis scaling
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.IsAutomaticMinValue = false;
        secondaryAxis.IsAutomaticMaxValue = false;
        secondaryAxis.IsAutomaticMajorUnit = false;
        secondaryAxis.MinValue = 0;
        secondaryAxis.MaxValue = 400;   // same max as primary axis
        secondaryAxis.MajorUnit = 100;  // same major unit as primary axis

        // Save the workbook
        workbook.Save("SecondaryAxisAligned.xlsx");
    }
}