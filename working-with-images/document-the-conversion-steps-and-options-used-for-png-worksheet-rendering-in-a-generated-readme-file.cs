// Title: Generate a high‑resolution transparent PNG from an Excel worksheet and create a README that logs the rendering steps using Aspose.Cells for .NET
// AI Prompts: Write C# code that builds a Workbook, fills cells with sample data, configures ImageOrPrintOptions for 300 DPI PNG with a transparent background, and uses SheetRender to save the first worksheet as a PNG file. | Add logic to the program that writes a README.txt file describing each conversion step, the chosen ImageOrPrintOptions values, and the generated output files. | Modify the rendering options to change the DPI, toggle the OnePagePerSheet flag, or disable transparency, and observe how the PNG output changes.
// Common Searches: asp.net c# how to export an Excel sheet to a transparent PNG with specific DPI using Aspose.Cells | example code for SheetRender ToImage with ImageOrPrintOptions one page per sheet | create README file programmatically after rendering Excel to image in C# | set horizontal and vertical resolution for PNG export in Aspose.Cells .NET | fit entire worksheet on one page when converting to PNG with Aspose.Cells
// Tags: worksheet PNG export using Aspose.Cells | ImageOrPrintOptions high DPI configuration | PNG rendering with alpha channel Aspose.Cells | SheetRender single-page PNG output | generate conversion README C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // This C# program creates a workbook with sample data, configures ImageOrPrintOptions for a 300 DPI PNG with a transparent background, renders the first worksheet to 'Worksheet.png' using SheetRender, and writes a README.txt that documents each step, the option values, and the resulting files.
class Program
{
    static void Main()
    {
        try
        {
            // Step 1: Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleSheet";
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.2);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.8);

            // Step 2: Define PNG rendering options (ImageFormat defaults to PNG)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Set the resolution (DPI)
                HorizontalResolution = 300,
                VerticalResolution = 300,
                // Fit the entire sheet on one page
                OnePagePerSheet = true,
                // Make background transparent (optional)
                Transparent = true
            };

            // Step 3: Render the worksheet to PNG
            SheetRender sr = new SheetRender(sheet, options);
            string pngPath = "Worksheet.png";
            sr.ToImage(0, pngPath);

            // Step 4: Generate README file documenting the conversion steps and options
            string readmePath = "README.txt";
            using (StreamWriter writer = new StreamWriter(readmePath))
            {
                writer.WriteLine("PNG Worksheet Rendering - Conversion Steps and Options");
                writer.WriteLine("----------------------------------------------------");
                writer.WriteLine();
                writer.WriteLine("1. Workbook Creation");
                writer.WriteLine("   - A new Workbook instance was created.");
                writer.WriteLine("   - Sample data was added to the first worksheet.");
                writer.WriteLine();
                writer.WriteLine("2. Rendering Options (ImageOrPrintOptions)");
                writer.WriteLine($"   - HorizontalResolution: {options.HorizontalResolution} DPI");
                writer.WriteLine($"   - VerticalResolution: {options.VerticalResolution} DPI");
                writer.WriteLine($"   - OnePagePerSheet: {options.OnePagePerSheet}");
                writer.WriteLine($"   - Transparent background: {options.Transparent}");
                writer.WriteLine();
                writer.WriteLine("3. Rendering Process");
                writer.WriteLine("   - SheetRender was instantiated with the target worksheet and options.");
                writer.WriteLine($"   - The first page of the sheet was exported to '{pngPath}' using ToImage(0, ...).");
                writer.WriteLine();
                writer.WriteLine("4. Output Files");
                writer.WriteLine($"   - {pngPath} : The rendered PNG image of the worksheet.");
                writer.WriteLine($"   - {readmePath}    : This documentation file.");
            }

            Console.WriteLine("Rendering completed. PNG and README files have been generated.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
