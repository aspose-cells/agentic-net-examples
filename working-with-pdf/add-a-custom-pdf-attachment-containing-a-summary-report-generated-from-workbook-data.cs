// Title: Create a PDF summary from Excel data and embed it as an OLE object using Aspose.Cells for .NET
// Common Searches: how to generate a PDF report from worksheet data and embed it in an Excel file using Aspose.Cells C# | Aspose.Cells C# add PDF as OLE object to existing .xlsx workbook | embed custom PDF attachment into Excel workbook programmatically with Aspose.Cells .NET | save PDF bytes as OLE object in Excel using Aspose.Cells API | calculate column total and attach summary PDF to workbook with Aspose.Cells
// Tags: PDF generation from Excel data Aspose.Cells | OLE object insertion of PDF in .xlsx | Workbook.Save with embedded PDF Aspose.Cells | OleObjects.Add method for PDF attachment | column sum calculation and PDF summary Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, computes the sum of the first column, creates a one‑cell PDF containing a summary report with that total, converts the PDF to a byte array, embeds it as an OLE object on the first worksheet, and saves the workbook with the PDF attachment.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook_WithPdfAttachment.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;

            // Calculate total of the first column (skip header row)
            int lastRow = cells.MaxDataRow;
            double total = 0;
            for (int row = 1; row <= lastRow; row++)
            {
                total += cells[row, 0].DoubleValue;
            }

            // Generate a PDF containing the summary using Aspose.Cells
            var pdfWorkbook = new Workbook();
            var pdfSheet = pdfWorkbook.Worksheets[0];
            pdfSheet.Cells[0, 0].PutValue($"Summary Report\r\nTotal of Column A: {total}");

            byte[] pdfBytes;
            using (var ms = new MemoryStream())
            {
                pdfWorkbook.Save(ms, SaveFormat.Pdf);
                pdfBytes = ms.ToArray();
            }

            // Embed the PDF as an OLE object in the original worksheet
            sheet.OleObjects.Add(0, 0, 100, 100, pdfBytes, "Summary.pdf");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
