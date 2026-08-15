// Title: C# – Apply "#,##0" number format to chart data labels with Aspose.Cells
// Description: Creates a workbook, fills a Category‑Quantity table, adds a column chart, enables data‑label values and sets the NumberFormat property to "#,##0" so quantities appear with thousand separators and no decimal places, then saves the file as NumberFormatChart.xlsx.
// Keywords: Aspose.Cells chart number format | C# thousand separator format | Excel chart data labels formatting | remove decimals from chart values | .NET Aspose.Cells custom format | NumberFormat property chart series
// Common Searches: Aspose.Cells set chart data label number format | C# format chart values with commas | How to hide decimals in Aspose.Cells chart labels | Apply custom number format to Excel chart using Aspose.Cells | Thousand separator for chart data labels .NET
// Developer Intent: Format chart data‑label numbers with commas and no decimal places using Aspose.Cells for .NET.
// Use Cases: Improve readability of large numeric values in column charts. | Generate Excel reports where chart labels show whole numbers only. | Maintain consistent numeric presentation across multiple charts programmatically.
// AI Prompts: Write C# code that creates a line chart with Aspose.Cells and applies the "#,##0" format to its data labels. | Show how to change the number format of pie‑chart data labels to include thousand separators using Aspose.Cells for .NET. | Provide an example of setting a custom number format for a chart series and exporting the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsNumberFormatDemo
{
    // Creates a workbook, fills a Category‑Quantity table, adds a column chart, enables data‑label values and sets the NumberFormat property to "#,##0" so quantities appear with thousand separators and no decimal places, then saves the file as NumberFormatChart.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["B2"].PutValue(1234);
            worksheet.Cells["B3"].PutValue(56789);
            worksheet.Cells["B4"].PutValue(101112);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and apply custom number format with thousand separators, no decimals
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.NumberFormat = "#,##0";

            // Save the workbook
            workbook.Save("NumberFormatChart.xlsx");
        }
    }
}
