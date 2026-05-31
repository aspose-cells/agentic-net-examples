using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class WaterfallSeriesToggle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a Waterfall chart
        // Column A – Category names, Column B – Values
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");

        string[] categories = { "Start", "Increase", "Decrease", "End" };
        double[] values = { 100, 30, -20, 110 };

        for (int i = 0; i < categories.Length; i++)
        {
            worksheet.Cells[i + 2, 0].PutValue(categories[i]); // A column
            worksheet.Cells[i + 2, 1].PutValue(values[i]);   // B column
        }

        // Add a Waterfall chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // ------------------------------------------------------------
        // Toggle visibility of individual series based on user preferences
        // ------------------------------------------------------------
        // Example user preferences: true = show series, false = hide series
        bool[] userPreferences = { true, false, true, false };

        // The IsFiltered property hides a series when set to true.
        // Therefore we assign the inverse of the user preference.
        for (int i = 0; i < chart.NSeries.Count && i < userPreferences.Length; i++)
        {
            chart.NSeries[i].IsFiltered = !userPreferences[i];
        }

        // Save the workbook with the updated chart
        workbook.Save("WaterfallSeriesToggle.xlsx");
    }
}