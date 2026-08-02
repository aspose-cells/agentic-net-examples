// Title: Toggle Chart Data Labels in Aspose.Cells for .NET (C#) Using a Boolean Flag
// Description: Creates a workbook, adds sample data and a column chart, then uses a configuration boolean to set DataLabels.ShowValue on the first series, enabling or disabling data labels before saving the file.
// Keywords: Aspose.Cells C# chart data labels | toggle data label visibility | DataLabels.ShowValue example | programmatic chart label control | Excel chart label flag
// Common Searches: Aspose.Cells hide chart data labels C# | set ShowValue flag for chart series Aspose.Cells | toggle data labels dynamically in .NET Excel | conditional chart label visibility Aspose.Cells
// Developer Intent: Enable or disable chart data labels at runtime based on a boolean configuration setting.
// Use Cases: Show labels only when a user selects an option in a generated report. | Produce cleaner charts for print or PDF export by hiding labels automatically. | Provide a single setting that controls label visibility across multiple series in a dashboard workbook.
// AI Prompts: How can I toggle data label visibility for all series in an Aspose.Cells chart using a config flag? | Show an example that reads a boolean from appsettings.json and applies it to DataLabels.ShowValue in C#. | Explain how to conditionally set ShowCategoryName and ShowValue for chart series based on user preferences.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelToggle
{
    // Creates a workbook, adds sample data and a column chart, then uses a configuration boolean to set DataLabels.ShowValue on the first series, enabling or disabling data labels before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Configuration flag to control data label visibility
            bool showDataLabels = true; // Set to false to hide data labels

            // Create a new workbook (lifecycle create)
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series and its DataLabels
            Series series = chart.NSeries[0];
            DataLabels dataLabels = series.DataLabels;

            // Toggle visibility of data labels based on the configuration flag
            dataLabels.ShowValue = showDataLabels;

            // Optionally, you can also control other label aspects, e.g. ShowCategoryName
            // dataLabels.ShowCategoryName = showDataLabels;

            // Save the workbook (lifecycle save)
            workbook.Save("ChartWithDataLabelsToggle.xlsx");
        }
    }
}
