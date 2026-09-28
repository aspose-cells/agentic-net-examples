// Title: Check for Chinese characters in an Excel chart title and export the chart as PNG using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, locates the first chart, validates that its Title.Text includes at least one Chinese Unicode character, and saves the chart as a PNG image. | Write a C# method that receives a workbook path, uses a \u4E00-\u9FFF regular expression to ensure the chart title contains Chinese characters, throws an exception if the check fails, and renders the chart to a PNG file via ImageOrPrintOptions.
// Common Searches: aspocells c# verify chart title contains Chinese characters before exporting to png | how to export a specific chart from an Excel workbook to PNG using Aspose.Cells | c# regex Unicode range \u4e00-\u9fff for chart title validation Aspose.Cells | detect non‑English characters in Aspose.Cells chart title and save as image
// Tags: chart title Chinese character check Aspose.Cells | render Excel chart as PNG using Aspose.Cells | Unicode range \u4e00-\u9fff regex for chart titles | image export settings for Aspose.Cells charts | ensure chart presence before image conversion Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System;
using System.IO;
using System.Text.RegularExpressions;

// Loads an .xlsx workbook, confirms a chart exists on the first worksheet, validates that the chart's title contains at least one Chinese character using a Unicode regex, and exports the chart to a PNG file with Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The file '{inputPath}' was not found.");

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure at least one chart is present
            if (sheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts found in the first worksheet.");

            Chart chart = sheet.Charts[0];

            // Verify the chart title contains Chinese characters
            string titleText = chart.Title.Text;
            if (!Regex.IsMatch(titleText, @"[\u4e00-\u9fff]"))
                throw new Exception("The chart title does not contain Chinese characters.");

            // Configure image export options (default format inferred from file extension)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Export the chart to a PNG file
            chart.ToImage(outputPath, imgOptions);

            Console.WriteLine("Chart title verified and PNG exported successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
