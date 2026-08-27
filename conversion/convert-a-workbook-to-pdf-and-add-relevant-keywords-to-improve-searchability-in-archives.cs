// Title: How to convert an Excel workbook to PDF and embed custom keyword metadata with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets the 'Keywords' built‑in document property, and saves the workbook as a PDF while preserving the document outline using Aspose.Cells. | Show how to configure PdfSaveOptions to enable ExportDocumentStructure and embed custom keywords before converting a workbook to PDF in C#.
// Common Searches: aspnet convert excel to pdf with custom pdf metadata aspose.cells | c# set built‑in document properties keywords before saving workbook as pdf | preserve outline exportdocumentstructure pdfsaveoptions aspose cells example
// Tags: Aspose.Cells set workbook keywords | PdfSaveOptions ExportDocumentStructure | Excel to PDF conversion with metadata | C# embed custom PDF metadata using Aspose | preserve document outline Aspose.Cells PDF

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfConversion
{
    // Loads an Excel workbook, assigns keyword values via BuiltInDocumentProperties, enables ExportDocumentStructure in PdfSaveOptions, and saves the workbook as a PDF.
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook
            // (Replace the path with the actual source file)
            Workbook workbook = new Workbook("input.xlsx");

            // Add keywords to the built‑in document properties.
            // These keywords become part of the PDF metadata after conversion,
            // improving searchability in archive systems.
            workbook.BuiltInDocumentProperties["Keywords"].Value = "Finance,Report,2023,Quarterly";

            // Create PDF save options.
            // ExportDocumentStructure = true preserves the document outline in the PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook as a PDF using the specified options.
            // (Replace the path with the desired output location)
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
