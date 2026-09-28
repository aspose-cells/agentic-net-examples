// Title: Flatten Excel comments and other annotations when converting a workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, calculates formulas, disables PDF document structure to flatten comments, and saves the workbook as a PDF using Aspose.Cells. | Show how to set PdfSaveOptions.ExportDocumentStructure to false in Aspose.Cells so that all worksheet annotations become part of the PDF page content. | Write a snippet that converts an Excel workbook to PDF while merging notes, comments, and shapes into the static PDF layers with Aspose.Cells.
// Common Searches: Aspose.Cells C# flatten Excel comments when exporting to PDF | disable PDF document structure in Aspose.Cells to merge annotations | convert XLSX to PDF with merged comments using Aspose.Cells | example of setting ExportDocumentStructure to false in Aspose.Cells PdfSaveOptions | export workbook to PDF without separate annotation objects Aspose.Cells
// Tags: flatten Excel annotations Aspose.Cells | PdfSaveOptions ExportDocumentStructure false | Excel to PDF conversion C# Aspose.Cells | merge worksheet comments into PDF page | pre-calculate formulas before PDF export Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfFlattenDemo
{
    // Loads an Excel workbook, optionally calculates formulas, disables PDF document structure to flatten all annotations, and saves the result as a PDF using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set ExportDocumentStructure to false so that annotations are flattened
            // and become part of the page content rather than separate PDF objects.
            pdfOptions.ExportDocumentStructure = false;

            // Optionally calculate formulas before saving
            workbook.CalculateFormula();

            // Save the workbook as a PDF using the configured options
            workbook.Save("output.pdf", pdfOptions);

            Console.WriteLine("Workbook has been converted to PDF with annotations flattened.");
        }
    }
}
