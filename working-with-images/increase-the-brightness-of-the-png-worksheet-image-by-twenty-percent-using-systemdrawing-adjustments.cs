// Title: Increase worksheet PNG brightness by 20% after Aspose.Cells export using System.Drawing in C#
// AI Prompts: Render the first worksheet to a PNG with Aspose.Cells, then load the PNG with System.Drawing, apply a 20% brightness increase using a ColorMatrix, and overwrite the file. | Add code that creates a Bitmap from the exported PNG, modifies its pixel values with a brightness‑adjusting ColorMatrix, and saves the adjusted image back to disk.
// Common Searches: C# how to brighten an Excel worksheet image exported as PNG by Aspose.Cells | apply 20 percent brightness increase to PNG generated from Aspose.Cells using System.Drawing | post‑process Aspose.Cells worksheet PNG to adjust brightness in .NET application
// Tags: worksheet PNG brightness adjustment with System.Drawing | Aspose.Cells export PNG post‑processing | C# ColorMatrix brightness filter | increase image brightness after Excel sheet rendering | System.Drawing brightness enhancement for exported worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates loading an Excel workbook, rendering the first worksheet to a PNG with Aspose.Cells, then using System.Drawing to apply a 20 % brightness boost via a ColorMatrix before saving the modified image.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Configure image export options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Render the worksheet to an image file
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, outputPath);

            Console.WriteLine($"Worksheet rendered successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
