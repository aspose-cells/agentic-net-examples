// Title: Overlay a PNG logo onto the first worksheet and export it as a branded PNG image using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, inserts a PNG logo at the calculated bottom‑right cell, and saves the sheet as a single PNG with 96 dpi using Aspose.Cells. | Write a script that checks for the workbook and logo files, adds the logo picture to the worksheet, configures ImageOrPrintOptions for PNG output, and renders the worksheet to a branded image.
// Common Searches: Aspose.Cells C# add picture to worksheet before rendering to PNG | how to place a logo in the bottom right corner of an Excel sheet and export as PNG using Aspose.Cells | render Excel worksheet as a single PNG image with custom logo Aspose.Cells .NET
// Tags: insert image into Excel worksheet Aspose.Cells | export worksheet as PNG with Aspose.Cells | configure ImageOrPrintOptions 96dpi | place logo at bottom right cell | render single-page PNG from Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, verifies the logo file, inserts a PNG logo at the bottom‑right corner of the first worksheet, configures ImageOrPrintOptions for 96 dpi PNG output, and renders the sheet to a single branded PNG image.
class Program
{
    static void Main()
    {
        // Input workbook and logo file paths
        string workbookPath = "input.xlsx";
        string logoPath = "logo.png";

        // Output path for the branded PNG image
        string outputPath = "branded_output.png";

        try
        {
            // Verify required files exist
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");
            if (!File.Exists(logoPath))
                throw new FileNotFoundException($"Logo file not found: {logoPath}");

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);
            Worksheet sheet = workbook.Worksheets[0]; // use the first worksheet

            // Determine approximate position for the logo (bottom‑right corner)
            int lastRow = sheet.Cells.MaxDataRow + 5;
            int lastColumn = sheet.Cells.MaxDataColumn + 5;

            // Add the logo picture to the worksheet; size is set automatically
            sheet.Pictures.Add(lastRow, lastColumn, logoPath);

            // Configure rendering options for PNG output
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,    // render the whole sheet as a single image
                HorizontalResolution = 96,
                VerticalResolution = 96
                // Default image format is PNG; no need to set explicitly
            };

            // Render the worksheet directly to a PNG file
            SheetRender sheetRender = new SheetRender(sheet, renderOptions);
            sheetRender.ToImage(0, outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
