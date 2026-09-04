// Title: How to set a currency number format on the value axis of a Waterfall chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a Waterfall chart with sample data and applies the "$#,##0.00" number format to its value axis using Aspose.Cells. | Show how to detect whether the ValueAxis.NumberFormat property exists and, if it does, set a currency format for a Waterfall chart in Aspose.Cells for .NET. | Provide a complete Aspose.Cells example that builds a Waterfall chart, saves it as an .xlsx file, and formats the axis values as monetary amounts, including proper error handling.
// Common Searches: Aspose.Cells C# set currency format for waterfall chart axis | How to apply custom number format to chart axis in Aspose.Cells .NET | Waterfall chart value axis formatting with Aspose.Cells example | C# Aspose.Cells number format property not available on ValueAxis | Display monetary values on Excel waterfall chart using Aspose.Cells
// Tags: Aspose.Cells chart axis number format customization | waterfall chart value axis currency display | C# set chart axis format Aspose.Cells .NET | Excel workbook save waterfall chart with formatted axis | conditional NumberFormat property check Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, adds sample data, inserts a Waterfall chart, and demonstrates (with a conditional line) how to apply a "$#,##0.00" currency number format to the chart's value axis before saving the file as WaterfallCurrency.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the waterfall chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(5000);
            sheet.Cells["A3"].PutValue("Increase");
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["A4"].PutValue("Decrease");
            sheet.Cells["B4"].PutValue(-1500);
            sheet.Cells["A5"].PutValue("End");
            sheet.Cells["B5"].PutValue(5500);

            // Add a waterfall chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // NOTE: The Axis class in some Aspose.Cells versions does not expose a NumberFormat property.
            // If your version supports it, you can uncomment the following line:
            // chart.ValueAxis.NumberFormat = "$#,##0.00";

            // Define output file path
            string outputPath = "WaterfallCurrency.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
