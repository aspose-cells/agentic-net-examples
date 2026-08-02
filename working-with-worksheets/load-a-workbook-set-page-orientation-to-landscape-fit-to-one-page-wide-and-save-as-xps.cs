// Title: C# – Export Excel to XPS with Landscape Orientation & Fit‑to‑Width using Aspose.Cells
// Description: This example shows how to load an .xlsx workbook, set the first worksheet to landscape mode, configure the page setup to fit the content to one page wide (height unrestricted), apply XpsSaveOptions for a single‑page output, and save the result as an XPS document with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# XPS export | landscape page orientation Aspose.Cells | fit to one page wide Excel | XpsSaveOptions OnePagePerSheet | Workbook.Save XPS | .NET Excel to XPS tutorial | page setup Aspose.Cells
// Common Searches: How to export Excel as XPS with landscape orientation in C# | Aspose.Cells fit worksheet to one page wide XPS | C# set page orientation to landscape before XPS export | XpsSaveOptions settings for single‑page Excel export | Save .xlsx to .xps using Aspose.Cells for .NET
// Developer Intent: Create an XPS file from an Excel workbook where the sheet prints in landscape mode and scales to a single page width.
// Use Cases: Generating printable XPS reports that require landscape layout and column visibility on one page. | Automating batch conversion of multiple worksheets to XPS while preserving a consistent page format. | Archiving Excel data as XPS with controlled orientation and scaling for cross‑platform viewing.
// AI Prompts: Write C# code with Aspose.Cells to load an .xlsx file, set landscape orientation, fit the sheet to one page wide, and save it as XPS using OnePagePerSheet. | Provide a C# loop that applies landscape orientation and fit‑to‑width settings to every worksheet in a workbook, then exports each to a separate XPS file. | Explain how to customize XpsSaveOptions (margins, image quality, compression) when exporting Excel to XPS with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example shows how to load an .xlsx workbook, set the first worksheet to landscape mode, configure the page setup to fit the content to one page wide (height unrestricted), apply XpsSaveOptions for a single‑page output, and save the result as an XPS document with Aspose.Cells for .NET.
class XpsExportExample
{
    static void Main()
    {
        // Load an existing workbook from file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (or iterate through all if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Set page orientation to Landscape
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Fit the worksheet to one page wide (height will adjust automatically)
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 0; // 0 means "as many pages as needed" for height

        // Prepare XPS save options (optional settings can be added here)
        XpsSaveOptions saveOptions = new XpsSaveOptions
        {
            // Ensure the whole sheet fits on a single page if desired
            OnePagePerSheet = true,
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook as XPS using the specified options
        string outputPath = "output.xps";
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine($"Workbook saved as XPS to '{outputPath}'.");
    }
}
