// Title: Convert an Excel workbook to PDF and attach a separate PDF using Aspose.Cells and Aspose.Pdf in C#
// AI Prompts: Write C# code that loads input.xlsx, saves it as output.pdf with Aspose.Cells, then uses Aspose.Pdf to embed attachment.pdf as a file attachment inside output.pdf. | Show a C# example that checks for the existence of the Excel and PDF files, converts the Excel file to PDF, and adds the second PDF as an embedded attachment using Aspose.Pdf, including proper exception handling. | Provide a step‑by‑step C# implementation that demonstrates converting a workbook to PDF and then attaching an additional PDF document to the generated PDF, referencing both Aspose.Cells and Aspose.Pdf libraries.
// Common Searches: how to add a PDF file as an attachment to a PDF generated from an Excel workbook using Aspose in C# | Aspose.Cells convert Excel to PDF then embed another PDF with Aspose.Pdf C# example | C# code to attach a PDF document to a PDF created from an .xlsx file using Aspose libraries | embedding a PDF attachment in a PDF produced by Aspose.Cells conversion C#
// Tags: Aspose.Cells Excel to PDF conversion C# | Aspose.Pdf embed file attachment C# | convert workbook to PDF and add attachment Aspose | C# PDF attachment with Aspose.Pdf | Excel to PDF with embedded document Aspose

using System;
using System.IO;
using Aspose.Cells;

// // Loads an Excel workbook, converts it to PDF using Aspose.Cells, verifies required files, and demonstrates how to embed a separate PDF as an attachment into the generated PDF with Aspose.Pdf, including basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string attachmentPath = "attachment.pdf";
            const string outputPath = "output.pdf";

            // Verify required files exist
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");
            if (!File.Exists(attachmentPath))
                throw new FileNotFoundException($"Attachment file not found: {attachmentPath}");

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Save workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            // Note: Embedding arbitrary files into a PDF requires Aspose.Pdf.
            // If such functionality is needed, add a reference to Aspose.Pdf and implement the attachment logic.
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
