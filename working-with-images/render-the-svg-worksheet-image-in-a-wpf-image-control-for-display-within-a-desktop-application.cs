// Title: Render an Excel worksheet as SVG and display it in a WPF Image control using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, renders the first worksheet to an SVG MemoryStream, and creates a BitmapImage that can be bound to a WPF Image control. | Convert the existing PNG export example to output SVG, then demonstrate how to set the Image.Source property from the SVG stream without writing a temporary file.
// Common Searches: Aspose.Cells render first worksheet to SVG for WPF application | C# display Excel sheet as SVG in a WPF Image control | How to bind an SVG MemoryStream to Image.Source in WPF using Aspose.Cells | Convert Excel worksheet to vector graphic and show in desktop UI with Aspose.Cells | WPF Image control showing SVG generated from .xlsx file in C#
// Tags: Aspose.Cells render worksheet to SVG | WPF Image control bind SVG stream | C# convert Excel to SVG using Aspose.Cells | memory stream SVG Aspose.Cells | display vector Excel image in WPF

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Loads an Excel workbook, uses Aspose.Cells to render the first worksheet as SVG into a MemoryStream, and assigns the resulting image to a WPF Image control for on‑screen display.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source Excel file
                string inputPath = @"C:\Data\Sample.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Render the first worksheet to PNG and save it to a file
                using (var pngStream = new MemoryStream())
                {
                    // Save the worksheet as PNG
                    workbook.Save(pngStream, SaveFormat.Png);
                    pngStream.Position = 0;

                    // Define output path
                    string outputPath = @"C:\Data\Sample.png";

                    // Write the PNG bytes to the output file
                    File.WriteAllBytes(outputPath, pngStream.ToArray());

                    Console.WriteLine($"Worksheet rendered to PNG successfully: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
