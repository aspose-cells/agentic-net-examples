// Title: Programmatically toggle data label visibility on an Aspose.Cells column chart using a configuration flag in C#
// AI Prompts: Write C# code that builds an Excel workbook, adds a column chart, and sets series.DataLabels.ShowValue based on a boolean read from a configuration source. | Demonstrate disabling category name, percentage, series name, and legend key for chart data labels when the visibility flag is false. | Show how to retrieve a setting from appsettings.json (or environment variables) and apply it to chart label visibility with Aspose.Cells.
// Common Searches: how to control Aspose.Cells chart data labels with a config setting in .NET | C# Aspose.Cells hide data labels conditionally based on appsettings | set series.DataLabels.ShowValue dynamically in Aspose.Cells workbook | disable chart label components Aspose.Cells column chart C# example
// Tags: Aspose.Cells conditional data label visibility | C# Aspose.Cells series.DataLabels.ShowValue | Aspose.Cells column chart label configuration | Excel workbook chart label toggle C# | Aspose.Cells read config for chart settings

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelToggle
{
    // The example creates a new workbook, populates sample data, adds a column chart, and uses a boolean configuration flag to control the series.DataLabels.ShowValue property. When the flag is false, it also turns off category name, percentage, series name, and legend key, then saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Configuration flag that determines whether data labels should be visible
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

            // Toggle data label visibility based on the configuration flag
            // ShowValue controls whether the value part of the data label is displayed.
            series.DataLabels.ShowValue = showDataLabels;

            // Optionally, hide other label components when disabling labels
            if (!showDataLabels)
            {
                series.DataLabels.ShowCategoryName = false;
                series.DataLabels.ShowPercentage = false;
                series.DataLabels.ShowSeriesName = false;
                series.DataLabels.ShowLegendKey = false;
            }

            // Save the workbook to a file
            workbook.Save("DataLabelToggleDemo.xlsx");
        }

        // Placeholder for retrieving the configuration flag.
        // In a real scenario this could read from appsettings, environment variable, etc.
        static bool GetShowDataLabelsFromConfig()
        {
            // For demonstration, toggle the flag here.
            // Set to true to show data labels, false to hide them.
            return true;
        }
    }
}
