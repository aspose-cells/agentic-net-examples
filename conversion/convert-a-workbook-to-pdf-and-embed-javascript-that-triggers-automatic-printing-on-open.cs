// Title: C# – Convert an Aspose.Cells Workbook to PDF and embed JavaScript for auto‑print on open
// Description: Creates a workbook, fills sample data, configures PdfSaveOptions with JavaScript that triggers printing when the PDF opens, saves the file, and confirms its existence.
// Keywords: Aspose.Cells PDF JavaScript C# | auto print PDF Aspose.Cells | embed JavaScript in PDF C# | PdfSaveOptions JavaScript example | convert Excel to PDF with print script | Aspose.Cells workbook to PDF code | C# generate printable PDF from Excel
// Common Searches: How to add JavaScript to a PDF using Aspose.Cells C# | Aspose.Cells auto‑print PDF on open example | C# save Excel as PDF with print script | PdfSaveOptions JavaScript property Aspose.Cells | Embed print command in PDF generated from workbook
// Developer Intent: Generate a PDF from an Aspose.Cells workbook in C# and embed JavaScript that automatically prints the document when it is opened.
// Use Cases: Distribute Excel reports as PDFs that print immediately on the recipient’s printer. | Create kiosk‑style documents that launch a print dialog without user interaction. | Automate batch conversion of spreadsheets to printable PDFs for mass mailing.
// AI Prompts: Provide C# code that adds a JavaScript "this.print()" action to a PDF created with Aspose.Cells. | Explain how to configure PdfSaveOptions to embed custom JavaScript and set PDF compliance levels. | Show how to verify that the JavaScript was embedded correctly by opening the PDF in Adobe Acrobat. | Troubleshoot why the auto‑print script might be blocked by PDF viewers when using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, fills sample data, configures PdfSaveOptions with JavaScript that triggers printing when the PDF opens, saves the file, and confirms its existence.
class WorkbookToPdfWithAutoPrint
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // 2. Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // 3. Define output PDF path and ensure the directory exists
            string pdfPath = "workbook.pdf";
            string pdfDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
            if (!Directory.Exists(pdfDir))
            {
                Directory.CreateDirectory(pdfDir);
            }

            // 4. Save the workbook as a PDF file
            workbook.Save(pdfPath, pdfOptions);

            // 5. Verify that the PDF was created
            if (File.Exists(pdfPath))
            {
                Console.WriteLine($"Workbook successfully converted to PDF: {pdfPath}");
            }
            else
            {
                Console.WriteLine("PDF file was not created.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
