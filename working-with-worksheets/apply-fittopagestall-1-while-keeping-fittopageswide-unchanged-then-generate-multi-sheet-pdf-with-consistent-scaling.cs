// Title: Aspose.Cells .NET: Export Multi‑Sheet PDF with FitToPagesTall = 1 (Width Unchanged)
// Description: Demonstrates how to create a workbook with two worksheets, define print areas, set PageSetup.FitToPagesTall = 1 while leaving FitToPagesWide at its default, and save the file as a single PDF using PdfSaveOptions (OnePagePerSheet = false, AllColumnsInOnePagePerSheet = false). The result is a multi‑sheet PDF that preserves vertical scaling on every sheet.
// Keywords: Aspose.Cells C# | FitToPagesTall | FitToPagesWide default | multi sheet PDF export | PdfSaveOptions OnePagePerSheet false | page setup scaling Aspose.Cells | vertical scaling PDF | Aspose.Cells .NET example | print area Aspose.Cells | PDF generation from workbook
// Common Searches: Aspose.Cells set FitToPagesTall for all worksheets | export workbook to PDF with consistent scaling | keep FitToPagesWide default when fitting to pages tall | multi‑sheet PDF Aspose.Cells .NET | PdfSaveOptions settings for vertical fit
// Developer Intent: Apply FitToPagesTall = 1 on each worksheet without altering FitToPagesWide, then generate a single PDF that contains all sheets with the same vertical scaling.
// Use Cases: Print reports where each sheet must fit on one page height but may extend horizontally. | Distribute large data tables that need only vertical scaling before merging into one PDF. | Create dashboard PDFs where sheet height is constrained while preserving original column widths.
// AI Prompts: Generate C# code that loops through every worksheet in an Aspose.Cells workbook, sets FitToPagesTall = 1 while leaving FitToPagesWide unchanged, and saves the workbook as a combined PDF. | Explain the impact of OnePagePerSheet and AllColumnsInOnePagePerSheet on PDF output when FitToPagesTall is used in Aspose.Cells. | Provide a step‑by‑step tutorial for configuring page setup to fit each sheet to one page tall and exporting all sheets to a single PDF with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsFitToPagesTallDemo
{
    // Demonstrates how to create a workbook with two worksheets, define print areas, set PageSetup.FitToPagesTall = 1 while leaving FitToPagesWide at its default, and save the file as a single PDF using PdfSaveOptions (OnePagePerSheet = false, AllColumnsInOnePagePerSheet = false). The result is a multi‑sheet PDF that preserves vertical scaling on every sheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sheet 1: populate with sample data
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            for (int i = 0; i < 120; i++)          // enough rows to span multiple pages
            {
                for (int j = 0; j < 8; j++)        // enough columns to span multiple pages
                {
                    sheet1.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Configure page setup for Sheet 1
            PageSetup ps1 = sheet1.PageSetup;
            ps1.PrintArea = "A1:H120";               // define the printable area
            ps1.FitToPagesTall = 1;                  // apply FitToPagesTall = 1
            // FitToPagesWide is left unchanged (default = 1)

            // -------------------------------------------------
            // Sheet 2: another sheet with different data
            // -------------------------------------------------
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            for (int i = 0; i < 80; i++)
            {
                for (int j = 0; j < 12; j++)
                {
                    sheet2.Cells[i, j].PutValue($"S{i + 1}C{j + 1}");
                }
            }

            // Configure page setup for Sheet 2 (same scaling settings)
            PageSetup ps2 = sheet2.PageSetup;
            ps2.PrintArea = "A1:L80";
            ps2.FitToPagesTall = 1;                  // same FitToPagesTall
            // FitToPagesWide remains default (1)

            // -------------------------------------------------
            // Save the workbook as a multi‑sheet PDF with consistent scaling
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure each sheet can span multiple pages (do not force one page per sheet)
                OnePagePerSheet = false,
                AllColumnsInOnePagePerSheet = false,
                // Keep the scaling defined by FitToPagesTall/Wide
                // No additional scaling updates required
            };

            // Save (lifecycle rule: save)
            workbook.Save("MultiSheetOutput.pdf", pdfOptions);

            Console.WriteLine("PDF generated with FitToPagesTall = 1 on all sheets.");
        }
    }
}
