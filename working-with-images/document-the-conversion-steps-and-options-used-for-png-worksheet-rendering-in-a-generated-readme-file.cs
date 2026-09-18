// Title: Convert Excel worksheets to PNG images with custom DPI and white background using Aspose.Cells for .NET and generate a README of the rendering steps
// AI Prompts: Write a C# console application that loads an .xlsx workbook, configures ImageOrPrintOptions (OnePagePerSheet = false, Transparent = false, 150 DPI), renders every worksheet page to separate PNG files, and creates a README.txt documenting the source file, option values, and output locations. | Adjust the PNG export code to produce a single combined PNG per worksheet by setting OnePagePerSheet = true, while still generating a README that records the applied rendering settings. | Add robust error handling to the rendering loop so that missing or corrupt worksheets are skipped and success or error messages are appended to the generated README summary.
// Common Searches: aspnet render excel sheet to png with specific dpi using aspose.cells | how to set transparent background to false when exporting Excel to PNG in C# | generate a readme file automatically after converting Excel worksheets to images | export each page of an Excel worksheet as separate PNG files with Aspose.Cells | configure ImageOrPrintOptions for multi-page PNG output in Aspose.Cells .NET
// Tags: Aspose.Cells ImageOrPrintOptions PNG export | C# render Excel worksheet to PNG | custom DPI setting Aspose.Cells | multiple page per sheet PNG output | auto-generate README after image conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPngRenderingDemo
{
    // // Loads an Excel workbook, sets ImageOrPrintOptions (OnePagePerSheet = false, Transparent = false, 150 DPI), renders each worksheet page to individual PNG files, and writes a README.txt that records the source file, rendering options, and output locations.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output paths
            string inputFile = "input.xlsx";               // Source workbook
            string outputFolder = "Output";                // Folder for PNG images
            string readmeFile = Path.Combine(outputFolder, "README.txt");

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            try
            {
                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputFile);

                // Configure rendering options for PNG output
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    // Render the entire sheet on a single page (false = multiple pages if needed)
                    OnePagePerSheet = false,

                    // Set background to white (transparent = false)
                    Transparent = false,

                    // Set the resolution (DPI) for the output image
                    HorizontalResolution = 150,
                    VerticalResolution = 150
                };

                // Iterate through each worksheet and render it to PNG
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];

                    // Create a SheetRender object for the current worksheet
                    SheetRender sheetRender = new SheetRender(sheet, renderOptions);

                    // Render each page of the worksheet (if multiple pages are generated)
                    for (int pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                    {
                        // Build the output PNG file name
                        string pngFileName = Path.Combine(outputFolder, $"{sheet.Name}_Page{pageIndex + 1}.png");

                        // Render the page to the PNG file (format inferred from file extension)
                        sheetRender.ToImage(pageIndex, pngFileName);
                    }
                }

                // Generate README file documenting the conversion steps and options used
                using (StreamWriter writer = new StreamWriter(readmeFile))
                {
                    writer.WriteLine("PNG Worksheet Rendering - Conversion Steps and Options");
                    writer.WriteLine("---------------------------------------------------");
                    writer.WriteLine();
                    writer.WriteLine("1. Load Workbook");
                    writer.WriteLine($"   - Source file: {inputFile}");
                    writer.WriteLine("   - API used: new Workbook(string fileName)");
                    writer.WriteLine();
                    writer.WriteLine("2. Configure ImageOrPrintOptions");
                    writer.WriteLine("   - OnePagePerSheet: false (allows multiple pages per sheet)");
                    writer.WriteLine("   - Transparent: false (white background)");
                    writer.WriteLine("   - HorizontalResolution: 150 DPI");
                    writer.WriteLine("   - VerticalResolution: 150 DPI");
                    writer.WriteLine();
                    writer.WriteLine("3. Render Each Worksheet");
                    writer.WriteLine("   - For each worksheet, a SheetRender object is created.");
                    writer.WriteLine("   - Each page of the worksheet is rendered to a separate PNG file.");
                    writer.WriteLine("   - File naming pattern: <WorksheetName>_Page<Number>.png");
                    writer.WriteLine();
                    writer.WriteLine("4. Output Files");
                    writer.WriteLine($"   - PNG images are saved in the \"{outputFolder}\" directory.");
                    writer.WriteLine($"   - README file (this document) is also saved in the same directory.");
                    writer.WriteLine();
                    writer.WriteLine("All steps are performed using Aspose.Cells for .NET APIs.");
                }

                Console.WriteLine("Rendering completed. PNG files and README.txt have been saved to the Output folder.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }
}
