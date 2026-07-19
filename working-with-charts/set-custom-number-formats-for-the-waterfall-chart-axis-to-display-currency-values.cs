// Title: Set Currency Number Format on Waterfall Chart Axis with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample financial data, inserts a Waterfall chart, and formats the value‑axis tick labels using a custom currency pattern (e.g., "$#,##0") via the TickLabels.NumberFormat property, then saves the Excel file.
// Keywords: Aspose.Cells | C# | Waterfall chart | currency number format | axis formatting | TickLabels | NumberFormat property | Excel automation | chart customization | financial reporting
// Common Searches: Aspose.Cells set chart axis currency format C# | Waterfall chart number format Aspose.Cells example | How to format value axis tick labels in Aspose.Cells | C# custom number format for Excel chart axis | Apply $ format to chart axis using Aspose.Cells
// Developer Intent: Apply a custom currency format to the value axis of a Waterfall chart.
// Use Cases: Generate financial statements where the waterfall chart displays amounts in dollars. | Build an executive dashboard that visualizes profit changes with currency‑styled axis labels. | Automate accounting worksheets that require consistent currency formatting across charts.
// AI Prompts: Show C# code to apply a "$#,##0" number format to a Waterfall chart axis with Aspose.Cells. | Give an example of using different currency symbols (e.g., €#,##0) for chart axes in Aspose.Cells for .NET. | Explain how TickLabels.NumberFormat affects the appearance of chart axis values in the saved Excel file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample financial data, inserts a Waterfall chart, and formats the value‑axis tick labels using a custom currency pattern (e.g., "$#,##0") via the TickLabels.NumberFormat property, then saves the Excel file.
    class WaterfallCurrencyAxisDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the waterfall chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Start");
                sheet.Cells["A3"].PutValue("Increase");
                sheet.Cells["A4"].PutValue("Decrease");
                sheet.Cells["A5"].PutValue("End");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(5000);
                sheet.Cells["B3"].PutValue(2000);
                sheet.Cells["B4"].PutValue(-1500);
                sheet.Cells["B5"].PutValue(5500);

                // Add a Waterfall chart
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the chart data range
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Apply custom currency number format to the value axis tick labels
                chart.ValueAxis.TickLabels.NumberFormat = "$#,##0";

                // Define output file path
                string outputPath = "WaterfallCurrencyAxis.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            WaterfallCurrencyAxisDemo.Run();
        }
    }
}
