// Title: C# – Export an Aspose.Cells workbook to PDF (Hello World example)
// Description: Creates a new Aspose.Cells workbook, writes "Hello" and "World" to cells A1 and B1, ensures the output folder exists, and saves the workbook as a PDF file (WorksheetImage.pdf) with full exception handling.
// Keywords: Aspose.Cells PDF export C# | save Excel as PDF .NET | Aspose.Cells workbook to PDF example | C# generate PDF from Excel data | Aspose.Cells SaveFormat.Pdf
// Common Searches: Aspose.Cells export worksheet to PDF C# | How to save Excel file as PDF using Aspose.Cells | C# code to convert workbook to PDF | Create PDF from Excel without Office Interop | Aspose.Cells SaveFormat.Pdf usage
// Developer Intent: Generate a PDF file directly from an Aspose.Cells workbook in a C# application.
// Use Cases: Automated reporting: convert generated Excel sheets to PDF for distribution. | Server‑side document conversion where Microsoft Office is unavailable. | Providing a downloadable PDF version of an Excel report in web or desktop apps.
// AI Prompts: Write C# code that builds an Aspose.Cells workbook, adds values to cells, and saves it as a PDF. | Show how to set page options (orientation, margins) when exporting an Aspose.Cells worksheet to PDF. | Explain best practices for handling file‑system errors and ensuring the output directory exists when saving PDFs with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Creates a new Aspose.Cells workbook, writes "Hello" and "World" to cells A1 and B1, ensures the output folder exists, and saves the workbook as a PDF file (WorksheetImage.pdf) with full exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // -------------------- Create and populate workbook --------------------
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];              // get the first worksheet
            sheet.Cells["A1"].PutValue("Hello");                   // sample data
            sheet.Cells["B1"].PutValue("World");

            // -------------------- Save workbook as PDF --------------------
            string pdfPath = "WorksheetImage.pdf";

            // Ensure the directory exists
            string pdfDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
            {
                Directory.CreateDirectory(pdfDir);
            }

            // Save the workbook directly to PDF format
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine("PDF created successfully at: " + Path.GetFullPath(pdfPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
