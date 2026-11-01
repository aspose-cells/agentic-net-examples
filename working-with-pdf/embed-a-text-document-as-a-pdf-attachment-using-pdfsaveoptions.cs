// Title: Convert an Excel workbook to PDF with Aspose.Cells in C#
// Common Searches: how to convert an .xlsx file to PDF using Aspose.Cells in C# | Aspose.Cells C# save workbook as PDF example | C# generate PDF from Excel workbook with Aspose.Cells library | using Aspose.Cells Workbook.Save to create PDF from Excel in .NET | convert Excel to PDF programmatically with Aspose.Cells C#
// Tags: Aspose.Cells Excel to PDF conversion C# | Workbook.Save PDF Aspose.Cells .NET | C# generate PDF from Excel workbook | Aspose.Cells SaveFormat.Pdf usage | programmatic Excel to PDF with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that 'input.xlsx' exists, loads it into an Aspose.Cells Workbook, and saves the workbook as 'output.pdf' using Workbook.Save with the PDF format, handling any exceptions that may occur.
class EmbedAttachmentExample
{
    static void Main()
    {
        try
        {
            // Verify the input Excel file exists
            string excelPath = "input.xlsx";
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input Excel file not found: {excelPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Define the output PDF path
            string pdfPath = "output.pdf";

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF saved successfully: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
