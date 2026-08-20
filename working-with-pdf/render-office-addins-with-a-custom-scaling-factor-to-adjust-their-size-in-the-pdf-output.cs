// Title: Render Office Add‑In with Custom Scaling When Saving to PDF using AspNet Aspose.Cells
// Description: Demonstrates how to set a worksheet's percent scaling (e.g., 150%) via PageSetup, verify the effective scale with SheetRender.PageScale, and export the workbook to PDF with PdfSaveOptions so the Add‑In UI appears larger.
// Keywords: Aspose.Cells PDF scaling | worksheet zoom Aspose.Cells | Office Add-In PDF export | SheetRender PageScale | PdfSaveOptions custom size | C# Aspose.Cells example
// Common Searches: Aspose.Cells set worksheet zoom before PDF export | how to apply percent scaling to PDF output in Aspose.Cells | retrieve page scale after setting PageSetup.Zoom | render Office Add‑In content to PDF with custom size
// Developer Intent: Apply a custom percent scaling factor to a worksheet that contains Office Add‑In content and generate a PDF that reflects the adjusted size.
// Use Cases: Increase worksheet scale to 150% so UI elements from an Office Add‑In are larger in the exported PDF. | Programmatically read the actual page scale after applying PageSetup.Zoom for logging or validation. | Create consistently sized PDF reports from Add‑In data across multiple pages using PdfSaveOptions.
// AI Prompts: Show C# code to set a worksheet's percent scaling and export it to PDF with Aspose.Cells. | How can I read the calculated page scale after applying PageSetup.Zoom in Aspose.Cells? | Explain the relationship between ImageOrPrintOptions, SheetRender, and PdfSaveOptions when scaling Office Add‑In content.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Saving;

namespace AsposeCellsAddInScalingDemo
{
    // Demonstrates how to set a worksheet's percent scaling (e.g., 150%) via PageSetup, verify the effective scale with SheetRender.PageScale, and export the workbook to PDF with PdfSaveOptions so the Add‑In UI appears larger.
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                RenderAddInWithCustomScale.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }

    public class RenderAddInWithCustomScale
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // 3. Populate some sample data (simulating an Office Add‑In content)
                for (int row = 0; row < 10; row++)
                {
                    sheet.Cells[row, 0].PutValue($"Item {row + 1}");
                    sheet.Cells[row, 1].PutValue((row + 1) * 10);
                }

                // 4. Set a custom scaling factor for printing/PDF output (150%)
                sheet.PageSetup.IsPercentScale = true;   // Use percent scaling
                sheet.PageSetup.Zoom = 150;              // 150% scaling

                // 5. Verify the calculated page scale using SheetRender
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();
                SheetRender renderer = new SheetRender(sheet, renderOptions);
                double calculatedScale = renderer.PageScale; // Reflects the 150% zoom
                Console.WriteLine($"Calculated page scale: {calculatedScale * 100}%");

                // 6. Prepare PDF save options (no extra scaling needed – the PageSetup.Zoom controls size)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // 7. Save the workbook as PDF with the custom scaling applied
                string pdfPath = "AddInScaledOutput.pdf";
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"Workbook saved to PDF with custom scaling at: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred in Run: {ex.Message}");
            }
        }
    }
}
