// Title: Toggle Excel chart data labels with a config flag – Aspose.Cells C# example
// Description: Creates a workbook, adds sample data, inserts a column chart, and uses a boolean configuration flag to show or hide data labels. When the flag is false, all label parts (value, category name, series name, percentage) are disabled before saving the file.
// Keywords: Aspose.Cells C# chart data labels | toggle data label visibility | Series.DataLabels.ShowValue | conditional chart labels | hide Excel chart labels programmatically
// Common Searches: how to hide chart data labels Aspose.Cells C# | toggle data labels with a setting in Aspose.Cells | conditional visibility of Excel chart labels .NET | Aspose.Cells Series.DataLabels example | programmatically control chart labels in C#
// Developer Intent: Enable or disable chart data labels at runtime based on a configuration setting.
// Use Cases: User‑driven reports where data labels can be turned on or off. | Generating clean printable charts by suppressing labels for PDF export. | Applying different label visibility rules for multiple series during automated workbook creation.
// AI Prompts: Generate C# code using Aspose.Cells that reads a boolean flag and toggles Series.DataLabels.ShowValue, also turning off ShowCategoryName, ShowSeriesName, and ShowPercentage when the flag is false. | Show how to load a setting from appsettings.json and apply it to chart label visibility in an Aspose.Cells workbook. | Explain how to iterate over all series in an Aspose.Cells chart and apply the same data‑label visibility logic based on a runtime configuration.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, and uses a boolean configuration flag to show or hide data labels. When the flag is false, all label parts (value, category name, series name, percentage) are disabled before saving the file.
class ToggleDataLabels
{
    static void Main()
    {
        // Configuration flag to control data label visibility
        bool showDataLabels = true; // Set to false to hide data labels

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

        // Toggle data label visibility based on the configuration flag
        series.DataLabels.ShowValue = showDataLabels;

        // When hiding labels, also ensure other label parts are disabled
        if (!showDataLabels)
        {
            series.DataLabels.ShowCategoryName = false;
            series.DataLabels.ShowSeriesName = false;
            series.DataLabels.ShowPercentage = false;
        }

        // Save the workbook to a file
        workbook.Save("ToggleDataLabelsDemo.xlsx");
    }
}
