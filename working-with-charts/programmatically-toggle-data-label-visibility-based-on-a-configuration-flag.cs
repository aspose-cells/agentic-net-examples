using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ToggleDataLabels
{
    static void Main()
    {
        // Configuration flag that determines whether data labels should be shown
        bool showDataLabels = GetShowDataLabelsFromConfig();

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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the first series of the chart
        Series series = chart.NSeries[0];

        // Toggle the visibility of data labels based on the configuration flag
        series.DataLabels.ShowValue = showDataLabels;

        // Optionally hide other label components
        series.DataLabels.ShowCategoryName = false;
        series.DataLabels.ShowSeriesName = false;

        // Save the workbook to a file
        workbook.Save("ToggleDataLabels.xlsx");
    }

    // Placeholder for retrieving the configuration flag.
    // Replace with actual configuration reading logic as needed.
    static bool GetShowDataLabelsFromConfig()
    {
        // Example: return true to show data labels, false to hide them.
        return true;
    }
}