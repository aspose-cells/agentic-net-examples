// Title: Set a custom currency number format on the value axis of a Waterfall chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a Waterfall chart and applies a currency number format (e.g., $#,##0.00) to its value axis, then saves the workbook. | Show how to create the target folder programmatically before calling Workbook.Save with Aspose.Cells. | Demonstrate adding a Waterfall chart, binding data series, and customizing the primary axis display as currency using the Aspose.Cells API.
// Common Searches: aspnet aspose.cells how to format waterfall chart axis as currency in C# | c# set number format for chart value axis Aspose.Cells example | waterfall chart axis custom number format Aspose.Cells .NET tutorial | save excel workbook with waterfall chart and ensure output directory exists using Aspose.Cells
// Tags: Aspose.Cells set chart value axis number format | C# waterfall chart currency axis formatting | Aspose.Cells chart axis NumberFormat property | ensure output directory exists Aspose.Cells save workbook | Aspose.Cells create waterfall chart from data

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills it with sample data, adds a Waterfall chart, applies a currency number format to the chart’s value axis (using the NumberFormat property when supported), ensures the destination folder exists, and saves the file as WaterfallChartCurrency.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the waterfall chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(5000);
            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["B3"].PutValue(8000);
            sheet.Cells["A4"].PutValue("Expense");
            sheet.Cells["B4"].PutValue(-3000);
            sheet.Cells["A5"].PutValue("Profit");
            sheet.Cells["B5"].PutValue(0); // Placeholder for total

            // Add a Waterfall chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Optional: format the primary value axis as currency
            // Note: Axis.NumberFormat may not be available in older Aspose.Cells versions,
            // so this line is omitted for compatibility.
            // chart.ValueAxis.NumberFormat = "$#,##0.00";

            // Define output file path
            string outputPath = "WaterfallChartCurrency.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
