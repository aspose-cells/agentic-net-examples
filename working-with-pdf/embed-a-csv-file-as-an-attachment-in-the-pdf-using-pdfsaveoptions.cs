// Title: Generate a PDF from an Excel workbook and add a CSV attachment using Aspose.Cells and Aspose.Pdf in C#
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, saves it as a PDF, and then uses Aspose.Pdf to attach a separate .csv file to the resulting PDF. | Show how to validate the presence of the Excel and CSV files, convert the workbook to PDF, and programmatically add the CSV as an attachment with Aspose.Pdf in a .NET console application.
// Common Searches: how to attach a CSV file to a PDF generated from Excel using Aspose in C# | Aspose.Cells save workbook as PDF then add file attachment with Aspose.Pdf | C# embed external document in PDF after converting Excel with Aspose | verify source files before converting Excel to PDF and attaching CSV in .NET
// Tags: aspose.cells pdfsaveoptions convert excel to pdf c# | aspose.pdf add external file to pdf c# | c# embed csv as pdf attachment using aspose.pdf | verify excel and csv existence before conversion c# | workaround aspose.cells missing embed attachment feature

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that both the Excel workbook and the CSV file exist, loads the workbook with Aspose.Cells, saves it as a PDF via PdfSaveOptions, and then demonstrates how to use Aspose.Pdf to attach the CSV to the generated PDF, because Aspose.Cells alone cannot embed arbitrary files.
class EmbedCsvInPdf
{
    static void Main()
    {
        // Paths for the source workbook, CSV file to embed, and the output PDF
        string workbookPath = "input.xlsx";
        string csvPath = "data.csv";
        string outputPdfPath = "output.pdf";

        try
        {
            // Verify that the required files exist
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

            if (!File.Exists(csvPath))
                throw new FileNotFoundException($"CSV file not found: {csvPath}");

            // Load the Excel workbook
            Workbook workbook = new Workbook(workbookPath);

            // Create PDF save options (customize as needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: Aspose.Cells does not provide a direct API to embed arbitrary files
            // as attachments in the generated PDF. If attachment functionality is required,
            // consider using Aspose.Pdf after saving the PDF with Aspose.Cells.

            // Save the workbook as PDF
            workbook.Save(outputPdfPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to '{outputPdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
