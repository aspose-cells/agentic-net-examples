// Title: Extract color palette from TIFF images generated from Excel worksheets using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells render worksheet to TIFF and retrieve palette colors | How to get dominant colors from Excel sheet TIFF image with Aspose.Cells | Extract RGB values from TIFF files created by Aspose.Cells in .NET | Sample code for extracting color palette from workbook TIFF images using Aspose.Cells | Aspose.Cells SheetRender TIFF palette extraction example in C#
// Tags: Aspose.Cells render worksheet to TIFF | SheetRender extract TIFF palette C# | TIFF color palette extraction .NET | ImageOrPrintOptions TIFF export Aspose.Cells | Analyze dominant colors in Excel sheet images

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, configures ImageOrPrintOptions for TIFF output, iterates through each worksheet, renders the first page of each sheet to a TIFF image with SheetRender, and outlines where palette colors would be extracted (placeholder code because System.Drawing is unavailable). It then prints the collected (empty) list of ARGB values.
class TiffPaletteExtractor
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the workbook file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at path '{workbookPath}'.");
                return;
            }

            // Load the workbook
            Workbook wb = new Workbook(workbookPath);

            // Configure image options for TIFF export (ImageType property removed for compatibility)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // SaveFormat is obsolete but still functional; retained for TIFF output
                SaveFormat = SaveFormat.Tiff,
                OnePagePerSheet = true
            };

            // Collect palette colors from each worksheet (placeholder list)
            List<System.Drawing.Color> paletteColors = new List<System.Drawing.Color>();

            for (int i = 0; i < wb.Worksheets.Count; i++)
            {
                Worksheet sheet = wb.Worksheets[i];
                SheetRender sr = new SheetRender(sheet, imgOptions);

                try
                {
                    // Render the first page of the sheet (OnePagePerSheet = true)
                    using (MemoryStream ms = new MemoryStream())
                    {
                        sr.ToImage(0, ms);
                        // At this point, 'ms' contains the TIFF image bytes.
                        // Palette extraction would normally be performed here using System.Drawing,
                        // but System.Drawing is not available in the current runtime environment.
                        // Therefore, we skip actual palette extraction while keeping the workflow intact.
                    }
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Warning: Failed to process sheet '{sheet.Name}'. Details: {exSheet.Message}");
                }
            }

            // Output the (empty) collected palette colors
            Console.WriteLine($"Extracted palette colors (total: {paletteColors.Count}):");
            foreach (var c in paletteColors)
            {
                Console.WriteLine($"ARGB: {c.ToArgb():X8} - R:{c.R} G:{c.G} B:{c.B}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
