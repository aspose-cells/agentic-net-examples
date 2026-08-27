// Title: Convert an Aspose.Cells Workbook to a High‑Quality XPS Document with One Page per Sheet in C#
// AI Prompts: Write C# code that creates a Workbook, fills cells with data, and saves it as an XPS file using XpsSaveOptions with OnePagePerSheet, DefaultFont set to Arial, and AllColumnsInOnePagePerSheet enabled. | Show how to configure page range, font compatibility checks, and column‑fitting settings when exporting an Excel workbook to XPS with Aspose.Cells Rendering.
// Common Searches: C# Aspose.Cells export workbook to XPS with one page per sheet | How to set default font for XPS output using Aspose.Cells Rendering | Aspose.Cells XpsSaveOptions page range and font compatibility example | Saving Excel as XPS for high‑quality printing in .NET | Aspose.Cells XPS conversion settings for fitting all columns on one page
// Tags: Aspose.Cells XPS export options | C# Excel to XPS conversion | single page per worksheet XPS output | Arial default font for XPS rendering | fit all columns on one XPS page

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsXpsConversion
{
    // The program creates a workbook, populates a few cells, configures XpsSaveOptions (one page per sheet, Arial as the default font, column fitting, page range, and font compatibility checks), and saves the workbook as a high‑quality XPS file named WorkbookDemo.xps.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells XPS Conversion Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(12345);
            sheet.Cells["B2"].PutValue(67.89);

            // Configure XPS save options (lifecycle: create)
            XpsSaveOptions saveOptions = new XpsSaveOptions
            {
                // Each sheet will be saved as a single page for high‑quality printing
                OnePagePerSheet = true,
                // Use a common font to ensure consistent rendering
                DefaultFont = "Arial",
                // Optional: specify page range (first page only)
                PageIndex = 0,
                PageCount = 1,
                // Ensure all columns fit on one page per sheet
                AllColumnsInOnePagePerSheet = true,
                // Enable font compatibility checks for better fidelity
                CheckFontCompatibility = true,
                CheckWorkbookDefaultFont = true
            };

            // Save the workbook as an XPS document using the save options (lifecycle: save)
            string outputPath = "WorkbookDemo.xps";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully saved as XPS to: {outputPath}");
        }
    }
}
