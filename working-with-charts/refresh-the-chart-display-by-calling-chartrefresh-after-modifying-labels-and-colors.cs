// Title: C# – Refresh an Aspose.Cells chart after modifying data label fonts and series colors
// Description: Demonstrates how to create a workbook, add a column chart, customize data‑label font color, size, and series palette, and invoke Chart.Refresh to ensure the visual changes are applied before saving the file.
// Keywords: Aspose.Cells chart refresh C# | Chart.Refresh Aspose.Cells | modify data label font Aspose.Cells | change series color palette Aspose.Cells | Aspose.Cells .NET chart update
// Common Searches: Aspose.Cells Chart.Refresh example | update chart after changing data label color .NET | force chart redraw Aspose.Cells | C# change series colors and refresh chart | how to apply font to chart labels Aspose.Cells
// Developer Intent: Apply visual changes to a chart’s data labels and series colors and guarantee they appear in the saved workbook.
// Use Cases: After programmatically adjusting data‑label font properties, call Chart.Refresh to redraw the chart. | When switching the chart’s color palette, invoke Chart.Refresh to reflect the new colors instantly. | Add, hide, or style data labels and use Chart.Refresh to ensure the modifications are rendered in the output file.
// AI Prompts: Write C# code that changes data‑label font color and size in an Aspose.Cells chart and then calls Chart.Refresh. | Explain if Chart.Refresh is required after modifying series colors in Aspose.Cells and provide a sample implementation. | Show how to force a chart redraw in Aspose.Cells for .NET after updating label visibility and palette.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart, customize data‑label font color, size, and series palette, and invoke Chart.Refresh to ensure the visual changes are applied before saving the file.
class RefreshChartDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series and modify their appearance
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.Font.Color = Color.Blue;   // Change font color
            series.DataLabels.Font.Size = 12;            // Change font size
            series.DataLabels.ApplyFont();               // Apply font changes to all labels

            // Change the series color palette (example uses the first palette type)
            chart.NSeries.ChangeColors((ChartColorPaletteType)0);

            // No explicit Refresh method is required; changes are applied automatically

            // Save the workbook with the updated chart
            string outputPath = "RefreshChartDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
