// Title: Delete rows 10‑15 in Excel and export to PDF with Aspose.Cells (C#)
// Description: C# example that loads an XLSX file with Aspose.Cells, removes rows 10‑15 from the first worksheet, and saves the cleaned workbook as a PDF using PdfSaveOptions. Includes file‑existence checks and automatic output folder creation.
// Keywords: Aspose.Cells | C# | .NET | delete rows Excel | remove rows 10‑15 | Excel to PDF conversion | PdfSaveOptions | worksheet row deletion | automated Excel processing | batch PDF export
// Common Searches: Aspose.Cells delete rows 10 to 15 C# | Export modified Excel sheet to PDF using Aspose.Cells | How to remove specific rows from an Excel workbook in .NET | C# code to delete rows and save as PDF with Aspose | PdfSaveOptions example for Excel to PDF conversion
// Developer Intent: Remove rows 10‑15 from an Excel worksheet and generate a PDF of the resulting sheet.
// Use Cases: Prepare a clean financial report by stripping out header rows before publishing as PDF. | Automate preprocessing of incoming Excel files: delete unwanted rows and deliver a PDF to end users. | Run a nightly batch job that validates source workbooks, removes rows 10‑15, and archives each as a PDF.
// AI Prompts: Write C# code that deletes rows 20‑30 from the second worksheet and saves the workbook as a landscape‑oriented PDF using Aspose.Cells. | Explain how to customize PdfSaveOptions (page size, margins, image quality) after deleting rows with Aspose.Cells. | Provide a step‑by‑step guide for handling missing input files, creating output directories, and logging errors when deleting rows and exporting to PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions

// C# example that loads an XLSX file with Aspose.Cells, removes rows 10‑15 from the first worksheet, and saves the cleaned workbook as a PDF using PdfSaveOptions. Includes file‑existence checks and automatic output folder creation.
class DeleteRowsAndExportPdf
{
    static void Main()
    {
        // Path to the source Excel file
        string inputFile = @"C:\Path\To\Your\SourceWorkbook.xlsx";

        // Path where the resulting PDF will be saved
        string outputPdf = @"C:\Path\To\Your\CleanedWorkbook.pdf";

        // Verify that the input file exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        // Ensure the output directory exists
        string outputDir = Path.GetDirectoryName(outputPdf);
        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (you can change the index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Delete rows 10 through 15 (zero‑based index: start at 9, delete 6 rows)
            worksheet.Cells.DeleteRows(9, 6);

            // Prepare PDF save options (optional – you can customize further if required)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the modified workbook as a PDF document
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine("Rows deleted and PDF saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
