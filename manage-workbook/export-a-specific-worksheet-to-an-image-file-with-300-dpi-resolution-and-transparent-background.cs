// Title: Export a single worksheet to a 300 DPI transparent PNG image using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, checks the file existence, and saves a chosen worksheet as a 300 DPI transparent PNG using Aspose.Cells. | Show how to configure ImageOrPrintOptions.HorizontalResolution, VerticalResolution, and Transparent properties to render a worksheet to a high‑resolution PNG in C#. | Write a robust method that catches exceptions while converting a specific sheet to a PNG image with custom DPI and transparency using Aspose.Cells.
// Common Searches: how to render an Excel worksheet to a 300‑dpi image with a transparent background in C# Aspose.Cells | Aspose.Cells export specific sheet as high resolution PNG image | C# set image resolution and enable transparency when saving a worksheet as an image | save first worksheet of workbook to PNG with 300 DPI using Aspose.Cells | image rendering options for transparent background in Aspose.Cells .NET
// Tags: worksheet to PNG conversion with custom DPI Aspose.Cells | transparent image rendering using ImageOrPrintOptions | C# Aspose.Cells set horizontal and vertical resolution | high‑resolution sheet image export using Aspose.Cells | file existence check before workbook load Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads input.xlsx, verifies the file, and renders the first worksheet to output.png as a 300 DPI PNG with a transparent background using ImageOrPrintOptions.
class ExportWorksheetToImage
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.png";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Select the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Configure image rendering options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300,
                Transparent = true
                // ImageFormat defaults to PNG; explicit setting omitted to avoid API mismatch
            };

            // Render the worksheet to an image file
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Worksheet exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
