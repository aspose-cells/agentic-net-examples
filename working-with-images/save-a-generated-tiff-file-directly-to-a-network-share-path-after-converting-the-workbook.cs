// Title: Render an Excel worksheet to a high‑resolution TIFF and save it directly to a UNC network share using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to render the first worksheet of a workbook as a 300 dpi single‑page TIFF and writes the file to a \\Server\Share UNC location. | Write a method that verifies a network folder exists (creating it if necessary) before calling SheetRender.ToImage to save the TIFF. | Show how to configure ImageOrPrintOptions for high‑resolution TIFF output and combine it with SheetRender to produce a one‑page image.
// Common Searches: Aspose.Cells export worksheet as TIFF to UNC path C# | how to save rendered TIFF image on a network share using Aspose.Cells | C# create missing network folder before saving Aspose.Cells image | set 300 dpi TIFF export options with Aspose.Cells ImageOrPrintOptions | render Excel sheet to single page TIFF and store on shared drive
// Tags: Aspose.Cells render worksheet to TIFF | high‑dpi TIFF export ImageOrPrintOptions | SheetRender ToImage remote path | pre‑save folder existence check C# | save rendered Excel image to remote folder

using System;
using System.IO;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Rendering;   // Required for ImageOrPrintOptions and SheetRender

// The example creates a workbook, fills it with sample data, configures the sheet to fit one page wide, sets high‑resolution (300 dpi) image rendering options, renders the first worksheet to a single‑page TIFF, ensures the target UNC directory exists (creating it if needed), and saves the TIFF directly to a network share.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(1800);

            // Fit the sheet on one page when printed
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;

            // Set image rendering options (default format will be used, e.g., PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 300,
                VerticalResolution = 300,
                OnePagePerSheet = true
            };

            // Render the worksheet to an image
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);

            // UNC network share path for the output file
            string networkPath = @"\\ServerName\ShareFolder\WorkbookOutput.tiff";

            // Ensure the target directory exists
            string directory = Path.GetDirectoryName(networkPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the rendered page to the network share
            sheetRender.ToImage(0, networkPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
