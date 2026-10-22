// Title: Convert an entire Excel workbook to a multi‑page TIFF using Aspose.Cells for .NET with default rendering settings
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and saves it as a multi‑page TIFF where each worksheet becomes a separate page, using the default SaveFormat.Tiff behavior. | Explain how to export all worksheets of a workbook to a single TIFF image in C# without customizing rendering options. | Show the minimal steps required to create a multi‑page TIFF from a workbook using Aspose.Cells' Workbook.Save method.
// Common Searches: Aspose.Cells C# save workbook as multi page TIFF default settings | how to export each Excel sheet to a separate page in a TIFF using Aspose.Cells | C# convert .xlsx to multi‑page TIFF without specifying image options | default rendering options for Workbook.Save to TIFF in Aspose.Cells
// Tags: Workbook.Save with SaveFormat.Tiff multi-page export | Aspose.Cells automatic TIFF rendering | C# export all worksheets to TIFF | multi-page TIFF generation from Excel | convert Excel workbook to TIFF image

using Aspose.Cells;
using System;

// Loads 'input.xlsx' into an Aspose.Cells Workbook and calls Workbook.Save with SaveFormat.Tiff, which automatically creates a multi‑page TIFF where each worksheet is rendered as a separate page, using the library's default rendering options.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Convert the entire workbook to a multi‑page TIFF using default rendering options
        // The Save method with SaveFormat.Tiff automatically creates a multi‑page TIFF
        // where each worksheet is rendered as a separate page.
        workbook.Save("output.tiff", SaveFormat.Tiff);
    }
}
