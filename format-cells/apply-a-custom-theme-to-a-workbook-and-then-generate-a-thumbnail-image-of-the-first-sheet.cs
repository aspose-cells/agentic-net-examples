// Title: Create a PNG thumbnail of the first worksheet after checking for a custom theme using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells generate PNG thumbnail of first worksheet | How to render an Excel sheet to an image with ImageOrPrintOptions in .NET | Check for customTheme.xml before saving workbook using Aspose.Cells | Aspose.Cells render sheet as single page PNG thumbnail | Save workbook after optional custom theme handling in C#
// Tags: Aspose.Cells render worksheet to PNG | ImageOrPrintOptions OnePagePerSheet example | SheetRender thumbnail generation C# | custom theme detection Aspose.Cells | save workbook after theme check .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a new Workbook, optionally detects a customTheme.xml file (noting that direct theme loading is unsupported), adds sample data, configures ImageOrPrintOptions with OnePagePerSheet, uses SheetRender to export the first worksheet as a PNG thumbnail, and saves both the thumbnail and the workbook while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Attempt to load a custom theme if the file exists
            const string themePath = "customTheme.xml";
            if (File.Exists(themePath))
            {
                // Direct theme loading is not available in the current Aspose.Cells .NET API.
                // Placeholder for custom theme handling if needed.
                Console.WriteLine($"Theme file found: {themePath}. Custom theme loading is not supported in this version.");
            }
            else
            {
                Console.WriteLine($"Theme file not found: {themePath}. Continuing without custom theme.");
            }

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B2"].PutValue(12345);

            // Set image options for the thumbnail
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true // Render the whole sheet as one page
            };

            // Render the first sheet to an image (thumbnail)
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            const string thumbnailPath = "FirstSheetThumbnail.png";
            renderer.ToImage(0, thumbnailPath);
            Console.WriteLine($"Thumbnail saved to {thumbnailPath}");

            // Save the workbook
            const string workbookPath = "WorkbookWithCustomTheme.xlsx";
            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved to {workbookPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
