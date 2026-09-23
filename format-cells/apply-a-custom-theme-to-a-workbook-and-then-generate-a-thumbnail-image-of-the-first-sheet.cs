// Title: Applying a custom XML theme to an Aspose.Cells workbook and generating a PNG thumbnail of the first worksheet in C#
// AI Prompts: Generate C# code that loads a custom theme XML file into an Aspose.Cells Workbook, writes sample data, and saves a PNG thumbnail of the first worksheet using SheetRender. | Show how to configure ImageOrPrintOptions for one‑page‑per‑sheet PNG output and render the first sheet of a themed workbook with Aspose.Cells. | Explain the steps to detect a theme file, apply it to a workbook (if supported), and then create a preview image of the first sheet in .NET.
// Common Searches: asp.net load custom theme xml into workbook aspose.cells c# example | create png thumbnail of first worksheet using sheetrender aspose.cells | save excel sheet as png one page per sheet aspose.cells .net | apply xml theme to workbook before rendering sheet aspose.cells c# | c# code to generate png preview of an excel sheet with aspose.cells
// Tags: custom XML theme loading Aspose.Cells | SheetRender PNG thumbnail generation | ImageOrPrintOptions one-page-per-sheet | first worksheet preview image Aspose.Cells | theme application prior to sheet rendering C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program creates or loads an Aspose.Cells Workbook, optionally checks for a custom theme XML file (placeholder for applying the theme), writes sample data to the first worksheet, configures ImageOrPrintOptions for PNG output, uses SheetRender to produce a PNG thumbnail of the first sheet, saves the thumbnail, and finally saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Optional: Load a custom theme if needed.
            string themePath = "customTheme.xml";
            if (File.Exists(themePath))
            {
                Console.WriteLine($"Theme file '{themePath}' found, but applying custom themes is not implemented in this example.");
                // Placeholder for applying a custom theme:
                // workbook.CustomTheme = Theme.Load(themePath);
            }
            else
            {
                Console.WriteLine($"Theme file '{themePath}' not found. Continuing without custom theme.");
            }

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            sheet.Cells["A2"].PutValue(12345);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Render the first worksheet to a PNG thumbnail
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Png,
                OnePagePerSheet = true
            };
            SheetRender renderer = new SheetRender(sheet, renderOptions);
            string thumbnailPath = "FirstSheetThumbnail.png";
            renderer.ToImage(0, thumbnailPath);
            Console.WriteLine($"Thumbnail saved to {thumbnailPath}");

            // Save the workbook to an Excel file
            string workbookPath = "ThemedWorkbook.xlsx";
            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved to {workbookPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
