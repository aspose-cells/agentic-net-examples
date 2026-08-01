// Title: Scale an Office Add‑In Worksheet for PDF Export with Aspose.Cells for .NET
// Description: C# example that creates a workbook, fills it with data, enables percent‑based scaling, sets a custom Zoom (e.g., 150%), optionally reads the effective PageScale via SheetRender, and saves the sheet to PDF using PdfSaveOptions. The resulting PDF reflects the specified scaling factor without altering the original layout.
// Keywords: Aspose.Cells PDF scaling | C# worksheet zoom export | PageSetup.Zoom Aspose.Cells | Office Add‑In PDF size | custom page scale PDF | SheetRender PageScale | PdfSaveOptions scaling | Aspose.Cells .NET example
// Common Searches: how to set worksheet zoom before PDF export Aspose.Cells | custom scaling factor for Office Add‑In PDF output C# | increase PDF page size using PageSetup.Zoom Aspose.Cells | retrieve actual page scale after zoom Aspose.Cells | export Excel sheet to PDF with 150% scaling
// Developer Intent: Apply a custom scaling factor to an Office Add‑In worksheet so that the exported PDF appears larger, using Aspose.Cells for .NET.
// Use Cases: Generate printable reports where charts and tables need to be enlarged without changing the source workbook. | Create PDF versions of Office Add‑In worksheets at a specific zoom level (e.g., 150% or 200%). | Validate the effective page scale before saving by reading SheetRender.PageScale.
// AI Prompts: Write C# code that sets a 200% zoom on a worksheet and exports it to PDF with Aspose.Cells. | Explain the impact of PageSetup.IsPercentScale and Zoom on PDF rendering in Aspose.Cells. | Show how to obtain the actual page scale after applying Zoom using SheetRender before PDF export.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Saving;

namespace AsposeCellsAddInPdfScalingDemo
{
    // C# example that creates a workbook, fills it with data, enables percent‑based scaling, sets a custom Zoom (e.g., 150%), optionally reads the effective PageScale via SheetRender, and saves the sheet to PDF using PdfSaveOptions. The resulting PDF reflects the specified scaling factor without altering the original layout.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            for (int row = 0; row < 20; row++)
            {
                sheet.Cells[row, 0].PutValue($"Item {row + 1}");
                sheet.Cells[row, 1].PutValue((row + 1) * 5);
            }

            // Define a custom scaling factor (e.g., 150% of original size)
            double customScaleFactor = 1.5; // 150%

            // Apply the scaling factor via PageSetup.Zoom (value is in percent)
            sheet.PageSetup.IsPercentScale = true;               // Ensure percent scaling is used
            sheet.PageSetup.Zoom = (int)(customScaleFactor * 100); // Set zoom to 150%

            // Optional: verify the calculated page scale using SheetRender
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();
            SheetRender renderer = new SheetRender(sheet, renderOptions);
            double calculatedScale = renderer.PageScale; // Should reflect the Zoom setting
            Console.WriteLine($"Calculated Page Scale after Zoom: {calculatedScale * 100}%");

            // Prepare PDF save options (no additional scaling needed here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook to PDF with the custom scaling applied
            string outputPath = "AddInScaledOutput.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF with custom scaling at: {outputPath}");
        }
    }
}
