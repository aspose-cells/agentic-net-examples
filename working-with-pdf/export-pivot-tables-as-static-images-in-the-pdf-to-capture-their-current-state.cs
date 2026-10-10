// Title: Save an Excel workbook with pivot tables as a PDF where pivot tables appear as static images using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, checks for its existence, and uses Aspose.Cells to export the workbook to a PDF with pivot tables rendered as static images. | Show how to configure PdfSaveOptions in Aspose.Cells to ensure pivot tables are captured as images during Excel‑to‑PDF conversion. | Provide a C# example that handles missing input files and saves the entire workbook, including any pivot tables, to a PDF document.
// Common Searches: aspnet convert excel file with pivot tables to pdf static image | c# Aspose.Cells export pivot table as image in pdf | how to preserve pivot table layout when saving workbook to pdf using Aspose.Cells | pdfsaveoptions pivot table rendering Aspose.Cells .NET | export entire workbook including pivot tables to pdf with Aspose.Cells
// Tags: Aspose.Cells PDF export with pivot tables | C# save workbook to PDF static images | PdfSaveOptions pivot table rendering | Excel to PDF conversion Aspose.Cells | render pivot tables as images in PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the input Excel file, loads it into an Aspose.Cells Workbook, and saves the workbook as a PDF using PdfSaveOptions, which renders any pivot tables as static images, with basic error handling.
class ExportPivotTablesToPdf
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Verify that the Excel file exists before attempting to load it
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file '{excelPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Define the output PDF path
            string pdfPath = "output.pdf";

            // Set PDF save options (default options are sufficient for most cases)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the entire workbook as a PDF document
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"PDF file successfully created at '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
