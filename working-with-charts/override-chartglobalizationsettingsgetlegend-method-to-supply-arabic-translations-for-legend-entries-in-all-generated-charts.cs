using Aspose.Cells;
using Aspose.Cells.Charts;

class ArabicChartGlobalizationSettings : ChartGlobalizationSettings
{
    // Arabic translation for legend increase label
    public override string GetLegendIncreaseName() => "زيادة";

    // Arabic translation for legend decrease label
    public override string GetLegendDecreaseName() => "نقصان";

    // Arabic translation for legend total label
    public override string GetLegendTotalName() => "المجموع";

    // Arabic translation for series name (used in legend)
    public override string GetSeriesName() => "السلسلة";

    // Arabic translation for "Other" label
    public override string GetOtherName() => "أخرى";

    // Optional: Arabic translation for chart title
    public override string GetChartTitleName() => "عنوان المخطط";

    // Optional: Arabic translation for axis title
    public override string GetAxisTitleName() => "عنوان المحور";
}

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("الفئة");
        worksheet.Cells["A2"].PutValue("أ");
        worksheet.Cells["A3"].PutValue("ب");
        worksheet.Cells["B1"].PutValue("القيمة");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";
        chart.Title.Text = "مثال المخطط";

        // Apply the Arabic globalization settings to all charts in the workbook
        workbook.Settings.GlobalizationSettings.ChartSettings = new ArabicChartGlobalizationSettings();

        // Save the workbook
        workbook.Save("ArabicChart.xlsx");
    }
}