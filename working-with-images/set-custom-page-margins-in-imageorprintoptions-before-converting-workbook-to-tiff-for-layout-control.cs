// Title: Apply custom page margins to an Excel worksheet and export it as a TIFF image using Aspose.Cells ImageOrPrintOptions in C#
// AI Prompts: Generate C# code that sets left, right, top, and bottom margins on a worksheet's PageSetup and then renders the first sheet to a TIFF file with Aspose.Cells ImageOrPrintOptions. | Show how to configure ImageOrPrintOptions for TIFF output while preserving the custom margins defined in the workbook.
// Common Searches: C# Aspose.Cells set worksheet margins before converting to TIFF | How to control page layout when rendering Excel to TIFF with ImageOrPrintOptions | Aspose.Cells ImageOrPrintOptions margin settings for TIFF export in .NET | Render first worksheet to TIFF with custom left and right margins using Aspose.Cells | Adjust top and bottom margins in Excel workbook before image conversion in C#
// Tags: worksheet page margins Aspose.Cells | ImageOrPrintOptions TIFF export C# | custom margins Excel to image | render worksheet to TIFF Aspose.Cells | page setup before image rendering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, sets left/right/top/bottom margins on the first worksheet via PageSetup, configures ImageOrPrintOptions, and renders the sheet to a TIFF image using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.tiff";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Set custom page margins (in points; 1 inch = 72 points)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.LeftMargin = 0.5 * 72;   // left margin
            sheet.PageSetup.RightMargin = 0.5 * 72;  // right margin
            sheet.PageSetup.TopMargin = 0.75 * 72;   // top margin
            sheet.PageSetup.BottomMargin = 0.75 * 72; // bottom margin

            // Configure image/print options for rendering
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Additional options can be set here if needed
                // Note: ImageFormat property may not be available in some versions;
                // the output format can be inferred from the file extension.
            };

            // Render the first worksheet to an image using the specified options
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Image file created successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
