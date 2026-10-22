// Title: Convert an Excel workbook to PDF while including hidden rows and columns using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export Excel to PDF with hidden rows and columns in C# | C# PdfSaveOptions to include hidden rows when saving workbook as PDF | How to keep hidden columns visible in PDF output using Aspose.Cells | Save Excel file as PDF preserving hidden data with Aspose.Cells .NET | Aspose.Cells PDF conversion include hidden rows example code
// Tags: Aspose.Cells PDF conversion include hidden rows | PdfSaveOptions hide rows Excel to PDF | C# export hidden columns to PDF Aspose.Cells | Excel to PDF preserving hidden data .NET | Aspose.Cells workbook save as PDF with hidden content

using System;
using System.IO;
using Aspose.Cells;

// The example creates a simple workbook, hides a specific row and column, then uses Aspose.Cells' PdfSaveOptions to save the workbook as a PDF that retains the hidden rows and columns in the output.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists; create a simple one if it does not.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample Data");
                ws.Cells["B2"].PutValue("Hidden Row Example");

                // Hide row 3 (index 2) and column C (index 2)
                ws.Cells.Rows[2].IsHidden = true;
                ws.Cells.Columns[2].IsHidden = true;

                wb.Save(inputPath);
            }

            // Load the workbook only if the file exists.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options.
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
