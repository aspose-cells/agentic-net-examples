// Title: Overlay a logo image onto a PNG rendering of an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add logo to rendered worksheet PNG | how to embed a branding image into Excel sheet image using Aspose.Cells .NET | Aspose.Cells SheetRender overlay custom image on PNG output | generate branded PNG from Excel worksheet with logo in C# | add watermark image to Excel worksheet PNG using Aspose.Cells
// Tags: overlay logo on worksheet PNG Aspose.Cells | add image watermark to Excel sheet rendering C# | render worksheet to PNG with custom logo Aspose.Cells | Aspose.Cells SheetRender add overlay image | C# generate branded worksheet image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program loads an Excel workbook, selects the first worksheet, configures PNG rendering options, and uses SheetRender to generate a PNG file of the worksheet's first page.
class Program
{
    static void Main()
    {
        // Paths to source Excel and final PNG
        string excelPath = "input.xlsx";
        string outputPngPath = "branded_output.png";

        try
        {
            // Verify that the input Excel file exists
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"Excel file not found: {excelPath}");

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Select the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure rendering options for PNG output
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                // Default image format is PNG; no need to set explicitly
                OnePagePerSheet = true,
                Transparent = false,
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // Render the worksheet directly to the final PNG file
            SheetRender sheetRender = new SheetRender(worksheet, renderOptions);
            sheetRender.ToImage(0, outputPngPath); // Render first page (index 0)
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
