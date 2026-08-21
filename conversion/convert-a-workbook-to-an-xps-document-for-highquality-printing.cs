// Title: C# – Convert an Aspose.Cells Workbook to XPS for High‑Quality Printing
// Description: Demonstrates how to create a workbook, add data, configure XpsSaveOptions (OnePagePerSheet, DefaultFont, page range) and save it as an XPS file using Aspose.Cells for .NET, ideal for crisp printable output.
// Keywords: Aspose.Cells | XPS conversion | .NET | C# | XpsSaveOptions | OnePagePerSheet | default font | high quality printing | Workbook.Save | export to XPS
// Common Searches: Aspose.Cells export workbook to XPS C# | How to set OnePagePerSheet in XpsSaveOptions | Save Excel as XPS with specific font using Aspose | Generate printable XPS from .NET workbook | Limit XPS output to first page Aspose.Cells
// Developer Intent: Create an XPS document from a workbook for precise, print‑ready output.
// Use Cases: Produce a single‑page XPS report for the first worksheet in an automated reporting pipeline. | Generate multi‑sheet XPS files where each sheet prints on its own page for batch printing. | Ensure consistent typography across printed XPS files by specifying a default font.
// AI Prompts: Show how to export every worksheet to separate XPS pages instead of only the first page. | Provide code that sets custom page size and margins for XPS conversion with XpsSaveOptions. | Explain how to embed a custom TrueType font in the XPS output and verify its presence.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsXpsConversion
{
    // Demonstrates how to create a workbook, add data, configure XpsSaveOptions (OnePagePerSheet, DefaultFont, page range) and save it as an XPS file using Aspose.Cells for .NET, ideal for crisp printable output.
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
                // Save only the first page (optional, can be omitted)
                PageIndex = 0,
                PageCount = 1
            };

            // Save the workbook as an XPS document (lifecycle: save)
            string outputPath = "WorkbookDemo.xps";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully saved as XPS to: {outputPath}");
        }
    }
}
