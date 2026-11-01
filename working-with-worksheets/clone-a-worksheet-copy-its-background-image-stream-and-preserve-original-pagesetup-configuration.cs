// Title: Clone an Excel worksheet, copy its background image, and retain page‑setup settings with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells clone worksheet and keep background image | Copy page setup (margins, print area, orientation) when duplicating a sheet using Aspose.Cells | Preserve print title rows and columns after AddCopy in Aspose.Cells | Transfer background image stream to a copied worksheet in .NET | How to duplicate a worksheet while retaining its page‑setup configuration with Aspose.Cells
// Tags: clone worksheet Aspose.Cells C# | copy worksheet background image Aspose.Cells | retain page setup properties Aspose.Cells | duplicate sheet preserve print area C# | transfer background image stream Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example loads an existing workbook, clones the first worksheet using AddCopy, assigns the source sheet's BackgroundImage to the cloned sheet, copies key PageSetup properties (paper size, orientation, fit‑to‑pages, print area, title rows/columns, and margins), and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            var workbook = new Workbook(inputPath);

            // Reference to the worksheet to be cloned (e.g., the first sheet)
            var sourceSheet = workbook.Worksheets[0];

            // Clone the worksheet; AddCopy returns the index of the new sheet
            int clonedIndex = workbook.Worksheets.AddCopy(sourceSheet.Index);
            var clonedSheet = workbook.Worksheets[clonedIndex];

            // ----- Copy background image -----
            if (sourceSheet.BackgroundImage != null)
            {
                // Directly assign the Image object to the cloned sheet
                clonedSheet.BackgroundImage = sourceSheet.BackgroundImage;
            }

            // ----- Preserve page‑setup configuration -----
            var srcSetup = sourceSheet.PageSetup;
            var dstSetup = clonedSheet.PageSetup;

            // Copy common page‑setup properties
            dstSetup.PaperSize = srcSetup.PaperSize;
            dstSetup.Orientation = srcSetup.Orientation;
            dstSetup.FitToPagesTall = srcSetup.FitToPagesTall;
            dstSetup.FitToPagesWide = srcSetup.FitToPagesWide;
            dstSetup.PrintArea = srcSetup.PrintArea;
            dstSetup.PrintTitleRows = srcSetup.PrintTitleRows;
            dstSetup.PrintTitleColumns = srcSetup.PrintTitleColumns;
            dstSetup.LeftMargin = srcSetup.LeftMargin;
            dstSetup.RightMargin = srcSetup.RightMargin;
            dstSetup.TopMargin = srcSetup.TopMargin;
            dstSetup.BottomMargin = srcSetup.BottomMargin;
            // Header and Footer copying omitted if not supported by the current API version

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
