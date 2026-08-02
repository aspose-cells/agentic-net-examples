// Title: C# – Export Excel to PDF with StandardSize optimization while preserving column widths (Aspose.Cells)
// Description: Demonstrates how to create a workbook, record its column widths, configure PdfSaveOptions with OptimizationType = Standard (StandardSize), and save the file as a PDF without altering the original column layout. Ideal for high‑print‑quality PDFs that retain exact column dimensions.
// Keywords: Aspose.Cells PDF StandardSize | preserve column widths | PdfSaveOptions OptimizationType Standard | C# export Excel to PDF | Aspose.Cells column width PDF | high quality PDF Aspose.Cells
// Common Searches: Aspose.Cells keep column widths when saving to PDF | StandardSize PDF optimization Aspose.Cells C# | PdfSaveOptions OptimizationType Standard example | Export Excel as PDF with original layout Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook using StandardSize optimization while ensuring the worksheet’s column widths remain unchanged.
// Use Cases: Produce print‑ready PDF reports that match the exact column layout defined in Excel. | Reduce PDF file size with StandardSize optimization without compromising column alignment. | Automate consistent PDF exports in batch processes where column dimensions must stay fixed.
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook to PDF with StandardSize optimization and retains original column widths. | Explain the impact of PdfSaveOptions.OptimizationType = PdfOptimizationType.Standard on PDF quality and layout. | Step‑by‑step guide to export Excel to PDF using Aspose.Cells while preserving column widths and applying StandardSize optimization.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfOptimizationDemo
{
    // Demonstrates how to create a workbook, record its column widths, configure PdfSaveOptions with OptimizationType = Standard (StandardSize), and save the file as a PDF without altering the original column layout. Ideal for high‑print‑quality PDFs that retain exact column dimensions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Description");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue("Fresh red apple");
            sheet.Cells["C2"].PutValue(1.25);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue("Ripe yellow banana");
            sheet.Cells["C3"].PutValue(0.80);

            // Preserve original column widths
            // Store current widths (optional, just for demonstration)
            double col0Width = sheet.Cells.GetColumnWidth(0);
            double col1Width = sheet.Cells.GetColumnWidth(1);
            double col2Width = sheet.Cells.GetColumnWidth(2);

            // Optionally set explicit widths to ensure they are not altered later
            sheet.Cells.SetColumnWidth(0, col0Width);
            sheet.Cells.SetColumnWidth(1, col1Width);
            sheet.Cells.SetColumnWidth(2, col2Width);

            // Create PDF save options and set the optimization type to Standard (high print quality)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OptimizationType = PdfOptimizationType.Standard   // StandardSize optimization
            };

            // Save the workbook as PDF while keeping the original column widths intact
            workbook.Save("Output_StandardSize.pdf", pdfOptions);

            Console.WriteLine("PDF saved with Standard optimization and original column widths preserved.");
        }
    }
}
