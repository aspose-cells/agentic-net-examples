using Aspose.Cells;
using Aspose.Cells.Charts;

// Custom globalization settings that localizes the "Other" label
class CustomChartGlobalizationSettings : ChartGlobalizationSettings
{
    // Override to provide a localized name for the "Other" slice in pie charts
    public override string GetOtherName()
    {
        return "Otros"; // Example: Spanish localization
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a pie chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["B3"].PutValue(50);
        worksheet.Cells["B4"].PutValue(20);

        // Add a pie chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Apply the custom globalization settings to the workbook
        GlobalizationSettings globalization = new GlobalizationSettings
        {
            ChartSettings = new CustomChartGlobalizationSettings()
        };
        workbook.Settings.GlobalizationSettings = globalization;

        // Save the workbook (the chart will use the localized "Other" label)
        workbook.Save("CustomOtherLabelPieChart.xlsx");
    }
}