// Title: Export a chart from an Excel workbook to a 300 DPI transparent PNG using Aspose.Cells for .NET
// AI Prompts: Create C# code that loads an .xlsx file, selects the first worksheet chart, and saves it as a 300 DPI PNG with a transparent background using Aspose.Cells. | Show how to configure ImageOrPrintOptions for high‑resolution PNG output when rendering a chart in Aspose.Cells. | Write a .NET snippet that checks for chart existence, renders it to a memory stream, and writes the PNG file to disk.
// Common Searches: Aspose.Cells C# export chart to 300 DPI PNG | How to save Excel chart as transparent PNG with Aspose.Cells | Render chart from workbook to high resolution image using .NET | Set DPI for chart image output in Aspose.Cells | Export first worksheet chart to PNG file programmatically
// Tags: chart.ToImage high DPI PNG Aspose.Cells | ImageOrPrintOptions transparent background C# | export Excel chart 300 DPI PNG | render workbook chart to memory stream Aspose | save chart image file with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// // Loads an Excel file, retrieves the first chart on the first worksheet, sets ImageOrPrintOptions to 300 DPI and transparent background, renders the chart to a memory stream, and saves it as a PNG file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart_high_res.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet and its first chart
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts were found in the first worksheet.");

            Chart chart = sheet.Charts[0];

            // Configure image options for high‑resolution PNG (300 DPI)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300, // DPI
                VerticalResolution = 300,   // DPI
                Transparent = true          // optional: keep background transparent
            };

            // Render the chart to a memory stream and save as PNG
            using (MemoryStream imgStream = new MemoryStream())
            {
                chart.ToImage(imgStream, imgOptions);
                File.WriteAllBytes(outputPath, imgStream.ToArray());
            }

            Console.WriteLine($"Chart image saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
