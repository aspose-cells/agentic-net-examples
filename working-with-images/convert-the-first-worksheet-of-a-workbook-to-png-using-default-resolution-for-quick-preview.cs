// Title: Generate a PNG preview of the first worksheet in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells render first worksheet to PNG image | How to create a quick PNG preview of the first sheet in an .xlsx with Aspose.Cells | Aspose.Cells SheetRender to PNG default resolution example in .NET | Saving first Excel worksheet as PNG using Aspose.Cells C# code | Render Excel worksheet page to PNG with default settings in Aspose.Cells
// Tags: sheetrender png export aspnet | first worksheet image rendering aspose.cells | excel to png preview using aspose.cells | default resolution png generation aspose.cells | c# workbook first sheet png conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample verifies the input file, loads the workbook, selects the first worksheet, applies default ImageOrPrintOptions (PNG format), uses SheetRender to render the first page, and saves it as a PNG preview file, handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "FirstSheetPreview.png";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet firstSheet = workbook.Worksheets[0];

            // Set image rendering options – default format is PNG
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Create a SheetRender object for the worksheet with the specified options
            SheetRender sheetRender = new SheetRender(firstSheet, imgOptions);

            // Render the first page (page index 0) of the worksheet to a PNG file
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Preview image saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
