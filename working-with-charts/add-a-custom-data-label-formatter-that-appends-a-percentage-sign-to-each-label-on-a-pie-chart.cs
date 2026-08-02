// Title: C# – Append a Percent Sign to Pie Chart Data Labels with Aspose.Cells
// Description: Creates a workbook, fills cells A1:B4 with categories and values, adds a pie chart, binds the series, enables data labels, shows only percentages, hides raw values, and applies the number format "0%" so each label ends with a percent sign before saving the file.
// Keywords: Aspose.Cells | C# | .NET | pie chart data labels | percentage format | custom number format | DataLabels.ShowPercentage | DataLabels.NumberFormat | Excel chart automation
// Common Searches: Aspose.Cells add percent sign to pie chart labels | C# set custom data label format for Excel chart | show only percentages on pie chart using Aspose.Cells | how to hide values and display percentages in chart labels .NET | custom number format for chart data labels Aspose
// Developer Intent: Configure a pie chart so each data label displays a formatted percentage with a trailing % sign.
// Use Cases: Generate Excel reports where pie chart labels show only percentages for clearer visualization. | Create automated dashboards that require custom‑formatted chart labels without raw numeric values. | Prepare financial summaries where chart labels must conform to a specific "0%" format.
// AI Prompts: Write C# code with Aspose.Cells to create a pie chart and set DataLabels to show percentages using the "0%" number format. | Show how to hide raw values and display only formatted percentages on a pie chart in .NET. | Explain the steps to apply a custom data label formatter that appends a percent sign to each label in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomDataLabelFormatter
{
    // Creates a workbook, fills cells A1:B4 with categories and values, adds a pie chart, binds the series, enables data labels, shows only percentages, hides raw values, and applies the number format "0%" so each label ends with a percent sign before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and configure custom formatting
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowPercentage = true;          // Show percentage values
            dataLabels.ShowValue = false;              // Hide raw values
            dataLabels.NumberFormat = "0%";            // Append a percent sign (custom formatter)

            // Save the workbook
            workbook.Save("PieChartWithCustomDataLabels.xlsx");
        }
    }
}
