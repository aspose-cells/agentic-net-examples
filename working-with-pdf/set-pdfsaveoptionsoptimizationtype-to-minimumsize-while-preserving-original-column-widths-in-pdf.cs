// Title: C# – Export Excel to PDF with MinimumSize optimization while preserving column widths using Aspose.Cells
// Description: Demonstrates how to set PdfSaveOptions.OptimizationType to MinimumSize and retain the worksheet's original column widths when saving an Excel workbook as a compact PDF with Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF export C# | PdfSaveOptions MinimumSize | preserve column widths PDF | Excel to PDF optimization | Aspose.Cells column width retention | C# PDF size reduction Aspose
// Common Searches: Aspose.Cells keep column widths when exporting to PDF | PdfSaveOptions OptimizationType MinimumSize C# example | How to reduce PDF size with Aspose.Cells without changing layout | Export Excel to PDF with fixed column widths Aspose
// Developer Intent: Configure PDF export to use MinimumSize optimization and ensure the original column widths are unchanged in the resulting PDF.
// Use Cases: Create lightweight PDF reports that match the Excel column layout for email distribution. | Generate PDF invoices where column alignment must stay consistent while minimizing file size. | Archive large spreadsheets as compact PDFs without losing visual formatting.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to PDF using MinimumSize optimization and keeps column widths intact. | Explain the effect of PdfSaveOptions.OptimizationType = MinimumSize on PDF file size and layout preservation. | Show how to set explicit column widths before exporting to PDF so they are retained with MinimumSize optimization.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to set PdfSaveOptions.OptimizationType to MinimumSize and retain the worksheet's original column widths when saving an Excel workbook as a compact PDF with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue("Data1");
        sheet.Cells["B2"].PutValue("Data2");

        // Set explicit column widths to ensure they are preserved in the PDF
        sheet.Cells.SetColumnWidth(0, 20); // Column A width
        sheet.Cells.SetColumnWidth(1, 30); // Column B width

        // Create PDF save options and set the optimization type to MinimumSize
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF while preserving the original column widths
        workbook.Save("PreservedColumns_MinSize.pdf", pdfOptions);
    }
}
