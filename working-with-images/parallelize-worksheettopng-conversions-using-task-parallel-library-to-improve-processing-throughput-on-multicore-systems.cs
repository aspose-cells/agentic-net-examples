// Title: Convert each worksheet of an Excel workbook to PNG in parallel using TPL and Aspose.Cells (C#)
// AI Prompts: Generate C# code that employs Parallel.For to render every worksheet of a workbook to a separate PNG file with Aspose.Cells, creating a new Workbook instance inside each iteration for thread safety. | Show how to configure ImageOrPrintOptions for PNG output and ensure the output directory is created automatically when processing worksheets in parallel. | Add per‑sheet exception handling that logs the sheet index and error while allowing the remaining worksheets to continue converting.
// Common Searches: how to speed up Excel worksheet to PNG conversion with Aspose.Cells and TPL | thread‑safe way to render multiple sheets to images in .NET | parallel processing of Excel sheets to PNG using Aspose.Cells C# example | best practice for multi‑core worksheet image export with Aspose.Cells
// Tags: TPL parallel worksheet image export Aspose.Cells | thread‑safe workbook instance per task C# | PNG rendering options Aspose.Cells | multi‑core Excel sheet conversion Aspose.Cells | parallel file output directory creation .NET

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program loads an Excel workbook, creates an output folder, and uses Parallel.For to render each worksheet to an individual PNG file with Aspose.Cells, instantiating a fresh Workbook inside each parallel iteration to maintain thread safety and handling errors per sheet.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook once (used only for sheet count)
            Workbook templateWorkbook = new Workbook(inputPath);
            int sheetCount = templateWorkbook.Worksheets.Count;

            // Ensure the output directory exists
            string outputDir = "output_png";
            Directory.CreateDirectory(outputDir);

            // Process each worksheet in parallel
            Parallel.For(0, sheetCount, sheetIndex =>
            {
                try
                {
                    // Load a fresh workbook instance for thread safety
                    Workbook wb = new Workbook(inputPath);

                    // Get the worksheet to render
                    Worksheet ws = wb.Worksheets[sheetIndex];

                    // Configure image rendering options (default format is PNG)
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        OnePagePerSheet = true
                    };

                    // Render the worksheet to PNG
                    SheetRender sr = new SheetRender(ws, imgOptions);
                    string outputPath = Path.Combine(outputDir, $"{ws.Name}.png");
                    sr.ToImage(0, outputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet index {sheetIndex}: {ex.Message}");
                }
            });

            Console.WriteLine("All worksheets have been converted to PNG.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
