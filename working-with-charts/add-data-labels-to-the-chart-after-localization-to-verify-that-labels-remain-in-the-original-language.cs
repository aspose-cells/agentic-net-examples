using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells; // for SettableChartGlobalizationSettings

class AddDataLabelsAfterLocalization
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Item A");
        sheet.Cells["A3"].PutValue("Item B");
        sheet.Cells["A4"].PutValue("Item C");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(250);
        sheet.Cells["B4"].PutValue(370);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // -------------------------------------------------
        // Localization step: customize globalizable strings
        // -------------------------------------------------
        // Create an instance of SettableChartGlobalizationSettings and set a localized name.
        // This simulates applying a different language (e.g., French) to chart elements.
        SettableChartGlobalizationSettings localization = new SettableChartGlobalizationSettings();
        localization.SetOtherName("Autre");               // "Other" label in French
        localization.SetLegendIncreaseName("Augmenter"); // "Increase" label in French
        // (In a real scenario the chart would consume these settings; here we just demonstrate usage.)

        // -------------------------------------------------
        // Add data labels after the localization step
        // -------------------------------------------------
        // Enable data labels for the first series and set custom text.
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;               // Show the numeric value
        series.DataLabels.ShowCategoryName = true;        // Show the category name

        // Verify that the label text remains in the original language (English in this example)
        // by explicitly setting it. The text should not be altered by the localization settings above.
        foreach (ChartPoint point in series.Points)
        {
            // Use the original English label format
            point.DataLabels.Text = $"Item: {point.XValue}, Value: {point.YValue}";
        }

        // Save the workbook to verify the result
        workbook.Save("DataLabelsAfterLocalization.xlsx");
    }
}