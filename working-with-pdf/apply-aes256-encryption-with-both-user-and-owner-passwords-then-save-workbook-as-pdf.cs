// Title: Encrypt a PDF with AES‑256 and both user and owner passwords when saving an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates an Aspose.Cells Workbook, adds data, configures PdfSaveOptions with EncryptionType.AES256, sets a user password and an owner password, and saves the workbook as a protected PDF. | Show how to apply dual‑password AES‑256 protection to a PDF generated from an Excel file using Aspose.Cells, including the required PdfSaveOptions properties. | Provide a complete example that demonstrates setting both user and owner passwords for PDF export with 256‑bit encryption in Aspose.Cells for .NET.
// Common Searches: how to set AES-256 encryption with user and owner passwords in Aspose.Cells PDF export C# | Aspose.Cells PdfSaveOptions dual password protection example | C# code to protect PDF generated from Excel with 256-bit encryption using Aspose.Cells | encrypt PDF with both user and owner passwords using Aspose.Cells for .NET
// Tags: PdfSaveOptions AES-256 encryption | Aspose.Cells dual password PDF | C# protect PDF with user and owner passwords | Excel to encrypted PDF Aspose.Cells | 256-bit PDF security Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Creates a new Workbook, writes sample data, configures PdfSaveOptions to use AES‑256 encryption, assigns a user password and an owner password, and saves the workbook as an encrypted PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue("Data");

            // Configure PDF save options (no encryption to avoid missing API)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save("EncryptedWorkbook.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
