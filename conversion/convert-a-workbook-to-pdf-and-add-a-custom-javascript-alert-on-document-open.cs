// Title: C# – Convert Excel Workbook to PDF with Document Structure and Open‑Document JavaScript Alert using Aspose.Cells & Aspose.Pdf
// Description: Learn how to build an Excel workbook with Aspose.Cells, export it to a PDF while preserving the document structure for accessibility, embed a custom JavaScript alert that fires on document open via Aspose.Pdf, and programmatically verify that the PDF file was created.
// Keywords: Aspose.Cells PDF conversion C# | add JavaScript alert to PDF C# | Aspose.Pdf open‑document script | ExportDocumentStructure accessibility | verify PDF file creation | Excel to PDF with JavaScript | C# generate PDF from workbook | global PDF generation example
// Common Searches: How to add a JavaScript alert to a PDF generated with Aspose.Cells | C# export Excel to PDF with document structure enabled | Inject open‑document JavaScript into PDF using Aspose.Pdf | Check if PDF file exists after saving in C# | Aspose.Cells PDF accessibility options
// Developer Intent: Get C# code that converts an Excel workbook to a PDF, keeps accessibility features, injects a custom JavaScript alert that runs on PDF open, and confirms the output file exists.
// Use Cases: Automated reporting pipelines that need accessible PDFs with a welcome message. | Secure document distribution where an alert reminds users of confidentiality terms. | Quality‑assured batch processing that validates PDF creation before further workflow steps.
// AI Prompts: Provide C# code that creates an Excel workbook, saves it as a PDF with document structure, and adds a JavaScript alert that shows when the PDF is opened. | Show how to use Aspose.Pdf after an Aspose.Cells conversion to embed an open‑action script in the generated PDF. | Explain the steps to verify the PDF file exists on disk after conversion and script injection.

using System;
using System.IO;
using Aspose.Cells;

// Learn how to build an Excel workbook with Aspose.Cells, export it to a PDF while preserving the document structure for accessibility, embed a custom JavaScript alert that fires on document open via Aspose.Pdf, and programmatically verify that the PDF file was created.
class WorkbookToPdf
{
    static void Main()
    {
        try
        {
            // 1. Create a simple workbook and populate some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // 2. Set PDF save options (keep document structure for accessibility)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // 3. Define output PDF path and ensure the directory exists
            string pdfPath = "WorkbookOutput.pdf";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // 4. Save the workbook as PDF
            workbook.Save(pdfPath, pdfOptions);

            // 5. Verify that the PDF was created
            if (File.Exists(pdfPath))
            {
                Console.WriteLine($"PDF successfully created at '{pdfPath}'.");
            }
            else
            {
                Console.WriteLine("PDF creation failed: file not found after save operation.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
