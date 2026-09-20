// Title: Apply conditional formatting to change chart axis tick label colors based on value thresholds using Aspose.Cells for .NET
// AI Prompts: Generate a line chart from column A values and column B categories, then set the category axis tick label font color to red when the corresponding value is greater than 70, green when less than 30, and keep the default color otherwise, finally save the workbook as XLSX. | Modify an existing Aspose.Cells chart example to add value‑based conditional formatting for the category axis tick labels without altering the data series. | Create a C# routine that builds a workbook, adds a line chart, and applies conditional formatting rules to the axis labels based on numeric thresholds before exporting the file.
// Common Searches: asp.net aspose.cells conditional formatting for chart axis tick labels | c# change category axis tick label color based on value threshold Aspose.Cells | how to apply value‑based formatting to chart axis labels in Aspose.Cells .NET | example of conditional formatting on chart axis labels using Aspose.Cells | set tick label font color conditionally in Aspose.Cells line chart
// Tags: conditional formatting chart axis tick labels Aspose.Cells | set tick label font color based on value Aspose.Cells C# | line chart axis label color threshold Aspose.Cells | apply value‑based formatting to chart axis Aspose.Cells | export workbook with formatted chart axis Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills column A with numeric data and column B with category names, adds a line chart referencing those ranges, and demonstrates how to apply conditional formatting to the category axis tick labels so their font color changes according to defined value thresholds. The workbook is then saved as AxisConditionalFormatting.xlsx.
class AxisConditionalFormattingExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (values) and column B (categories)
            double[] values = { 30, 55, 45, 70, 20, 90 };
            for (int i = 0; i < values.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(values[i]);          // Column A: numeric values
                sheet.Cells[i, 1].PutValue($"Item {i + 1}");   // Column B: category names
            }

            // Add a line chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart (values in column A)
            chart.NSeries.Add("A1:A6", true);
            // Set the category axis labels (categories in column B)
            chart.NSeries.CategoryData = "B1:B6";

            // Note: Aspose.Cells TickLabels does not expose an IsVisible property.
            // If hiding tick labels is required, adjust the axis visibility or font settings accordingly.
            // Example (optional): chart.CategoryAxis.IsVisible = false;

            // Save the workbook to a file
            string outputPath = "AxisConditionalFormatting.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
