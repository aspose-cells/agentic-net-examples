// Title: Aspose.Cells C# – Turn Off Chart Data Label Text Wrapping (DataLabels.IsTextWrapped = false)
// Description: C# example that creates a workbook, adds a column chart with sales data, shows values on the first series' data labels, disables text wrapping via DataLabels.IsTextWrapped, and saves the file as ChartDataLabels_NoWrap.xlsx.
// Keywords: Aspose.Cells chart data label wrap | DataLabels.IsTextWrapped false | C# Aspose.Cells disable label wrap | Excel chart label formatting .NET | Aspose.Cells chart label options
// Common Searches: Aspose.Cells disable data label wrapping | DataLabels.IsTextWrapped C# example | turn off text wrap for chart labels Aspose.Cells | chart data label formatting Aspose.Cells .NET | how to prevent label line breaks in Aspose.Cells chart
// Developer Intent: Disable text wrapping on chart data labels in Aspose.Cells.
// Use Cases: Generate a column chart where numeric labels stay on a single line for a clean visual layout. | Create Excel reports with long category names but keep data labels unwrapped for readability. | Automate workbook creation for dashboards that require precise label formatting without line breaks.
// AI Prompts: Write C# code using Aspose.Cells to add a bar chart and set DataLabels.IsTextWrapped = false for all series. | Explain the effect of DataLabels.IsTextWrapped and how to apply it to multiple series in a chart. | Show how to combine DataLabels.IsTextWrapped with other label settings such as ShowValue and ShowCategoryName in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDataLabelWrapDemo
{
    // C# example that creates a workbook, adds a column chart with sales data, shows values on the first series' data labels, disables text wrapping via DataLabels.IsTextWrapped, and saves the file as ChartDataLabels_NoWrap.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(15000);
            worksheet.Cells["B3"].PutValue(18000);
            worksheet.Cells["B4"].PutValue(21000);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the DataLabels of the first series
            DataLabels dataLabels = chart.NSeries[0].DataLabels;

            // Show the values on the data labels
            dataLabels.ShowValue = true;

            // Disable text wrapping for the data labels
            dataLabels.IsTextWrapped = false;

            // Save the workbook to a file
            workbook.Save("ChartDataLabels_NoWrap.xlsx");
        }
    }
}
