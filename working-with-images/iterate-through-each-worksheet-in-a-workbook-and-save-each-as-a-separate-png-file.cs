// Title: Save Each Excel Worksheet as a PNG Image Using Aspose.Cells for .NET
// Description: This example loads an Excel workbook, sets ImageOrPrintOptions to generate one page per sheet, creates an output directory, loops through all worksheets, and uses SheetRender to export the first page of each sheet to a uniquely named PNG file. Includes basic file‑existence checks and error handling.
// Keywords: Aspose.Cells PNG export | C# convert Excel sheet to image | ImageOrPrintOptions OnePagePerSheet | SheetRender save worksheet as PNG | batch export Excel worksheets to PNG | .NET Excel to PNG conversion
// Common Searches: how to export each Excel sheet as PNG in C# | Aspose.Cells render multiple worksheets to images | save workbook worksheets as separate PNG files | C# code to convert Excel worksheets to PNG using Aspose | batch image export of Excel sheets .NET
// Developer Intent: Generate an individual PNG file for every worksheet in a workbook.
// Use Cases: Create thumbnail previews of each sheet for a web‑based file explorer. | Archive workbook content as image files for compliance or documentation. | Supply separate chart images to reporting dashboards that require sheet‑level visuals.
// AI Prompts: Write C# code that opens an Excel file, iterates through its worksheets, and saves each one as a PNG with Aspose.Cells, handling missing files and rendering errors. | Explain how ImageOrPrintOptions.OnePagePerSheet and SheetRender collaborate to produce one PNG per worksheet. | Modify the sample so that each worksheet is saved in its own subfolder named after the sheet, preserving the original file name.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // This example loads an Excel workbook, sets ImageOrPrintOptions to generate one page per sheet, creates an output directory, loops through all worksheets, and uses SheetRender to export the first page of each sheet to a uniquely named PNG file. Includes basic file‑existence checks and error handling.
    public class SaveWorksheetsAsPng
    {
        public static void Run()
        {
            try
            {
                // Path to the input workbook
                string workbookPath = "input.xlsx";

                // Verify that the input file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Input file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Configure image rendering options: one page per sheet
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true
                    // ImageFormat defaults to PNG; explicit setting omitted to avoid compatibility issues
                };

                // Ensure the output directory exists
                string outputDir = "output";
                Directory.CreateDirectory(outputDir);

                // Iterate through each worksheet and render it to a separate PNG file
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];

                    // Create a SheetRender for the current worksheet with the specified options
                    SheetRender sheetRender = new SheetRender(sheet, options);
                    try
                    {
                        // Build a filename that includes the sheet index and name
                        string fileName = $"Sheet_{i + 1}_{sheet.Name}.png";
                        string filePath = Path.Combine(outputDir, fileName);

                        // Render the first (and only) page of the sheet to the PNG file
                        sheetRender.ToImage(0, filePath);

                        Console.WriteLine($"Worksheet '{sheet.Name}' saved as PNG: {filePath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to render sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                SaveWorksheetsAsPng.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
