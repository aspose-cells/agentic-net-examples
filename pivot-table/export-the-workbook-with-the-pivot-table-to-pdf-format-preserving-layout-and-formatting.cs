// Title: Export an Excel workbook with a pivot table to PDF while preserving column widths and page breaks using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file containing a pivot table and saves it as a PDF with original column widths and multi‑page layout using Aspose.Cells. | Show how to set Aspose.Cells PdfSaveOptions to disable column scaling and one‑page‑per‑sheet behavior when converting a workbook with pivot tables to PDF.
// Common Searches: asp.net aspose.cells export pivot table to pdf keep original layout | c# pdfsaveoptions allcolumnsinonepageper sheet false example | how to preserve column widths when converting Excel to PDF with Aspose.Cells | export workbook with pivot tables to multi‑page PDF using Aspose.Cells .NET
// Tags: Aspose.Cells PDF export preserving layout | PdfSaveOptions column width retention | export pivot table to PDF .NET | Excel to PDF conversion with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input Excel file, loads it with Aspose.Cells, configures PdfSaveOptions to keep original column widths and allow the sheet to span multiple pages, then saves the workbook—including any pivot tables—as a PDF while handling exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the pivot table.
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to keep the original layout and formatting.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not force all columns onto a single page; keep original column widths.
                AllColumnsInOnePagePerSheet = false,
                // Allow the sheet to span multiple pages if necessary.
                OnePagePerSheet = false
                // Additional options such as grid lines, headings, or compliance can be set
                // if supported by the current Aspose.Cells version.
            };

            // Export the workbook (including the pivot table) to PDF.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
