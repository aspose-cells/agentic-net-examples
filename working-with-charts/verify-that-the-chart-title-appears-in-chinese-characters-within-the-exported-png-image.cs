// Title: Export an Excel chart to PNG and ensure its title includes Chinese characters with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, checks the first chart's title for Chinese characters using a regular expression, and saves the chart as a PNG with Aspose.Cells. | Create a method that validates a chart title contains CJK Unified Ideographs before calling Chart.ToImage and ImageOrPrintOptions. | Generate a script that confirms the exported PNG file exists and has a non‑zero size after rendering a chart from a workbook.
// Common Searches: Aspose.Cells C# export chart as PNG and verify Chinese title | how to check Excel chart title for CJK characters before image export | using ImageOrPrintOptions to render chart to PNG in .NET | validate PNG file size after exporting chart with Aspose.Cells | regex pattern for Chinese characters in Excel chart title C#
// Tags: export chart to png Aspose.Cells | validate chart title Chinese characters C# | regex cjk detection Aspose.Cells | imageorprintoptions chart rendering .NET | check png file size after export

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System;
using System.IO;
using System.Text.RegularExpressions;

// Loads an Excel workbook, ensures the first worksheet contains a chart, verifies the chart title includes Chinese characters via a CJK regex, exports the chart to a PNG image using ImageOrPrintOptions, and confirms the PNG file was created and is not empty.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputImagePath = "chart.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one chart
            if (worksheet.Charts.Count == 0)
                throw new Exception("No charts found on the first worksheet.");

            Chart chart = worksheet.Charts[0];

            // Retrieve the chart title (may be null)
            string chartTitle = chart.Title?.Text ?? string.Empty;

            // Verify that the title contains Chinese characters (CJK Unified Ideographs)
            if (!Regex.IsMatch(chartTitle, @"\p{IsCJKUnifiedIdeographs}+"))
                throw new Exception("The chart title does not contain Chinese characters.");

            // Set image export options (format inferred from file extension)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Export the chart to a PNG image
            chart.ToImage(outputImagePath, imgOptions);

            // Validate that the PNG file was created and is not empty
            if (!File.Exists(outputImagePath))
                throw new Exception("The exported PNG image was not created.");

            FileInfo fi = new FileInfo(outputImagePath);
            if (fi.Length == 0)
                throw new Exception("The exported PNG image is empty.");

            Console.WriteLine("Chart title verified (contains Chinese characters) and PNG exported successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
