// Title: Create a PNG preview of the first worksheet in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file, verifies its existence, and uses Aspose.Cells SheetRender with default ImageOrPrintOptions to save the first worksheet as a PNG image. | Show how to handle a missing input workbook and ensure the output directory is created while converting the first sheet to a PNG preview with Aspose.Cells. | Demonstrate rendering a single worksheet to PNG at the default resolution using the Aspose.Cells Rendering API in a .NET console application.
// Common Searches: Aspose.Cells C# export first worksheet to PNG without specifying resolution | How to generate a quick PNG preview of an Excel sheet using Aspose.Cells .NET | C# sample for converting the first sheet of an .xlsx file to a PNG image with default settings | Render Excel worksheet to image with Aspose.Cells handling missing file errors | SheetRender ToImage example for single worksheet PNG output in .NET
// Tags: Aspose.Cells SheetRender export worksheet to PNG | C# default image rendering options Aspose.Cells | convert first Excel sheet to PNG preview | handle missing workbook file Aspose.Cells | create output directory before saving image Aspose.Cells | ImageOrPrintOptions default resolution PNG

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example checks that the source Excel file exists, loads it with Aspose.Cells, retrieves the first worksheet, applies default ImageOrPrintOptions (PNG format, default resolution), renders the sheet using SheetRender, and saves the result as a PNG file while handling directory creation and potential errors.
class WorksheetToPng
{
    static void Main()
    {
        // Path to the source workbook
        string inputFile = "input.xlsx";

        // Path for the output PNG image
        string outputFile = "sheet1.png";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file '{inputFile}' not found.");
            return;
        }

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputFile);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Get the first worksheet (index 0)
            Worksheet firstSheet = workbook.Worksheets[0];

            // Set up image rendering options (default resolution, PNG format)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            // The default ImageFormat is PNG, so no explicit setting is required.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                try
                {
                    Directory.CreateDirectory(outputDir);
                }
                catch (Exception dirEx)
                {
                    Console.WriteLine($"Failed to create output directory: {dirEx.Message}");
                    return;
                }
            }

            // Render the worksheet to an image (lifecycle rule: create/save handled by Aspose)
            SheetRender renderer = new SheetRender(firstSheet, imgOptions);
            renderer.ToImage(0, outputFile); // Render page 0 (the only page) to PNG

            Console.WriteLine("Worksheet converted to PNG successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}
