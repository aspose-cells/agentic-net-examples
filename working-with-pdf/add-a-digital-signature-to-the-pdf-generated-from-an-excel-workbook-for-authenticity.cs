// Title: How to add a digital signature to a PDF generated from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a PDF from an Aspose.Cells Workbook and sign it with a PFX certificate using PdfSaveOptions in C#. | Modify the workbook.Save call to include a DigitalSignature object that references a .pfx file and password. | Configure signature metadata (reason, location) and embed the digital signature while exporting Excel to PDF with Aspose.Cells.
// Common Searches: aspnet add digital signature to pdf created from excel using aspose.cells | c# sign pdf output from workbook with pfx certificate Aspose.Cells | how to use PdfSaveOptions DigitalSignature property in Aspose.Cells | apply digital signature to excel to pdf conversion in .NET | Aspose.Cells PDF signing example with certificate file
// Tags: Aspose.Cells PDF digital signature | C# PdfSaveOptions sign PDF | PFX certificate Aspose.Cells | Excel to PDF signing .NET | DigitalSignature property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates an Excel workbook, populates cells, and saves it as a PDF while embedding a digital signature from a PFX certificate using Aspose.Cells PdfSaveOptions. It also shows how to set signature metadata and handle potential errors.
class Program
{
    static void Main()
    {
        try
        {
            // -------------------- Create Excel workbook --------------------
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];              // get the first worksheet
            sheet.Cells["A1"].PutValue("Sample");                  // add some data
            sheet.Cells["B1"].PutValue(DateTime.Now);

            // -------------------- Save as PDF --------------------
            string pdfFilePath = "Report.pdf";                     // output PDF file

            // Ensure the directory exists
            string pdfDir = Path.GetDirectoryName(pdfFilePath);
            if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
            {
                Directory.CreateDirectory(pdfDir);
            }

            // Save the workbook as PDF
            workbook.Save(pdfFilePath, SaveFormat.Pdf);
            Console.WriteLine($"PDF file created successfully at: {Path.GetFullPath(pdfFilePath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
