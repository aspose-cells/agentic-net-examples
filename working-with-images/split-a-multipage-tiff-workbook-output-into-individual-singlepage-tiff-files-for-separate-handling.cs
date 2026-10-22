// Title: How to split a multi‑page TIFF generated from an Excel worksheet into separate single‑page TIFF files using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, renders each printed page of a worksheet, and saves each page as an individual TIFF file. | Demonstrate using SheetRender and ImageOrPrintOptions to export every worksheet page to separate TIFF images with sequential filenames. | Create a console application that iterates over SheetRender.PageCount and calls ToImage for each page to produce Page_1.tiff, Page_2.tiff, etc.
// Common Searches: Aspose.Cells C# split multi page TIFF into separate files per worksheet page | How to export each printed page of an Excel sheet as individual TIFF images using Aspose.Cells | C# code to render Excel worksheet pages to separate TIFF files with Aspose.Cells | Save multi‑page TIFF from Excel as separate page files using Aspose.Cells .NET
// Tags: Aspose.Cells SheetRender export single-page TIFF | C# Aspose.Cells render worksheet pages to TIFF | ImageOrPrintOptions SaveFormat TIFF multi-page split | Aspose.Cells split multi-page TIFF files | Excel worksheet to separate TIFF pages Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, configures ImageOrPrintOptions for TIFF, uses SheetRender to determine page count, then iterates through each page and saves it as a separate TIFF file named Page_1.tiff, Page_2.tiff, and so on.
class TiffSplitter
{
    static void Main()
    {
        // Load the source workbook (the workbook that will be rendered to a multi‑page TIFF)
        // Replace "source.xlsx" with the path to your original Excel file.
        Workbook workbook = new Workbook("source.xlsx");

        // Choose the worksheet you want to split. Here we use the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Configure rendering options for TIFF output.
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
        {
            // Set the output format to TIFF.
            SaveFormat = SaveFormat.Tiff,
            // Each page will be rendered separately; we do not combine pages into one image.
            OnePagePerSheet = false
        };

        // Create a SheetRender object which provides page‑by‑page rendering capabilities.
        SheetRender sheetRender = new SheetRender(sheet, renderOptions);

        // Get the total number of pages that the worksheet will occupy when printed.
        int pageCount = sheetRender.PageCount;

        // Iterate through each page and save it as an individual TIFF file.
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
        {
            // Build a file name for the current page (e.g., "Page_1.tiff", "Page_2.tiff", ...).
            string outputFileName = $"Page_{pageIndex + 1}.tiff";

            // Render the specific page to the file.
            sheetRender.ToImage(pageIndex, outputFileName);
        }

        Console.WriteLine("TIFF pages have been split into separate files successfully.");
    }
}
