using Aspose.Cells;
using Aspose.Cells.Charts;

class SetOtherLabelDemo
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
        worksheet.Cells["A5"].PutValue("D");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);
        worksheet.Cells["B5"].PutValue(40);

        // Add a pie chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Pie, 7, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B5", true);               // Values
        chart.NSeries.CategoryData = "A2:A5";           // Categories
        chart.Title.Text = "Sample Pie Chart";

        // Create a SettableChartGlobalizationSettings instance and set custom "Other" label
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetOtherName("Miscellaneous Items");

        // Apply the chart globalization settings to the workbook
        GlobalizationSettings globalization = new GlobalizationSettings();
        globalization.ChartSettings = chartSettings;
        workbook.Settings.GlobalizationSettings = globalization;

        // Save the workbook
        workbook.Save("PieChartOtherLabel.xlsx");
    }
}