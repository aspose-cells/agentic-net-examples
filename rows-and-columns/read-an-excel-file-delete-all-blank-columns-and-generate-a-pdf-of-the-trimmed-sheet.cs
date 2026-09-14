// Title: Remove blank columns from an Excel worksheet and save the cleaned sheet as a PDF with Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file, delete every empty column in the first worksheet, and export the result to a single‑page PDF using Aspose.Cells in C#. | Trim the worksheet by removing all blank columns, then set the PDF options so each sheet is rendered on one page, and save the workbook as PDF with Aspose.Cells. | Clean up an Excel sheet programmatically, eliminate any columns without data, and convert the cleaned workbook to a one‑page‑per‑sheet PDF via Aspose.Cells.
// Common Searches: C# Aspose.Cells delete empty columns before PDF conversion | How to remove blank columns from an Excel file and export to PDF using .NET | Aspose.Cells remove empty columns example with PDF conversion | Generate a one‑page PDF from an Excel worksheet after trimming blank columns in C# | Remove all blank columns in first worksheet and save as PDF Aspose.Cells
// Tags: DeleteBlankColumns Aspose.Cells | PdfSaveOptions OnePagePerSheet C# | blank column removal Excel Aspose.Cells | Excel to PDF conversion after column cleanup | trim worksheet before PDF export Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTrimAndExportPdf
{
    // Loads an Excel workbook, deletes all blank columns from the first worksheet using Aspose.Cells, then saves the trimmed workbook as a single‑page PDF with appropriate PdfSaveOptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path
            string inputFile = "input.xlsx";

            // Output PDF file path
            string outputFile = "trimmed.pdf";

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (you can change the index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Delete all blank columns in the worksheet
            worksheet.Cells.DeleteBlankColumns();

            // Create PDF save options (optional configuration)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure each sheet is saved on a single page (adjust as required)
                OnePagePerSheet = true
            };

            // Save the trimmed workbook as a PDF file
            workbook.Save(outputFile, pdfOptions);

            Console.WriteLine($"Workbook processed and saved to PDF: {outputFile}");
        }
    }
}
