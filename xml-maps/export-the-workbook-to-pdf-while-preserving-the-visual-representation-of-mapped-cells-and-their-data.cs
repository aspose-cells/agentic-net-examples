// Title: Export Workbook to PDF with Preserved XML Map Layout using AspNet.Cells C#
// Description: Demonstrates how to save an Aspose.Cells workbook as a PDF while keeping the visual appearance of XML‑mapped cells by enabling the ExportDocumentStructure option and optionally calculating formulas.
// Keywords: Aspose.Cells PDF export C# | PdfSaveOptions ExportDocumentStructure | preserve XML map layout PDF | workbook to PDF with visual structure | C# Aspose.Cells export mapped cells | calculate formulas before PDF save | Aspose.Cells PDF conversion example
// Common Searches: Aspose.Cells keep XML map visual when saving to PDF | PdfSaveOptions ExportDocumentStructure C# example | export workbook to PDF preserving cell layout | how to calculate formulas before PDF export Aspose.Cells | C# code to convert spreadsheet with XML map to PDF
// Developer Intent: Generate a PDF from a workbook that retains the exact visual layout of cells linked to an XML map.
// Use Cases: Create printable reports from spreadsheets that contain XML‑mapped data without losing the mapping view. | Produce compliance‑ready PDFs where calculated formulas and cell positioning must match the original worksheet. | Automate batch conversion of multiple mapped workbooks to PDFs while preserving each document’s visual fidelity.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook with an XML map to PDF, ensuring the visual structure is retained. | Explain the impact of PdfSaveOptions.ExportDocumentStructure on the resulting PDF and when to enable it. | Show how to trigger formula calculation before saving a workbook as PDF with Aspose.Cells in C#.

using System;
using Aspose.Cells;

namespace AsposeCellsPdfExportDemo
{
    // Demonstrates how to save an Aspose.Cells workbook as a PDF while keeping the visual appearance of XML‑mapped cells by enabling the ExportDocumentStructure option and optionally calculating formulas.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (including an XML map scenario if needed)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Create PDF save options and enable document structure export
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true   // Preserve visual representation of mapped cells
            };

            // Ensure formulas are calculated before saving (optional but recommended)
            workbook.CalculateFormula();

            // Save the workbook to PDF using the provided Save method (lifecycle rule)
            workbook.Save("MappedWorkbook.pdf", pdfOptions);

            Console.WriteLine("Workbook exported to PDF with document structure preserved.");
        }
    }
}
