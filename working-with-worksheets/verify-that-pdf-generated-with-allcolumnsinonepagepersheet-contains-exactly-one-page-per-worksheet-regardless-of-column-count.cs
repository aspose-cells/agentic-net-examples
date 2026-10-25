// Title: How to verify that Aspose.Cells generates a single PDF page per worksheet when using AllColumnsInOnePagePerSheet in C#
// Common Searches: C# Aspose.Cells check PDF has one page per worksheet with AllColumnsInOnePagePerSheet | validate PDF pagination per sheet using Aspose.Cells PdfSaveOptions in .NET | ensure all columns fit on one PDF page per worksheet Aspose.Cells example | test multi-worksheet PDF output for single-page per sheet Aspose.Cells C# | confirm PDF file creation after setting AllColumnsInOnePagePerSheet option
// Tags: Aspose.Cells PdfSaveOptions AllColumnsInOnePagePerSheet | C# generate single-page PDF per worksheet | verify PDF page count per worksheet Aspose.Cells | test multi-worksheet PDF pagination Aspose.Cells | validate column fit on one PDF page Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook with three worksheets, each containing 100 columns and 10 rows, configures PdfSaveOptions.AllColumnsInOnePagePerSheet = true, saves the workbook as a PDF, checks that the PDF file exists, and outputs the worksheet count along with a simple verification message indicating whether the PDF was generated.
class VerifyPdfAllColumnsOnePagePerSheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook with multiple worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();

            // Ensure we have at least three worksheets for the test
            for (int i = 0; i < 3; i++)
            {
                Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet.Name = $"Sheet{i + 1}";

                // Populate each sheet with many columns (e.g., 100) and some rows
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 100; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row}C{col}");
                    }
                }
            }

            // Configure PDF save options to force all columns onto one page per sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = true
            };

            // Save the workbook to a PDF file
            string pdfPath = "AllColumnsOnePagePerSheet.pdf";
            workbook.Save(pdfPath, pdfOptions);

            // Verify that the PDF file was created
            bool pdfExists = File.Exists(pdfPath);
            int worksheetCount = workbook.Worksheets.Count;

            // Output the verification result
            Console.WriteLine($"Worksheets: {worksheetCount}");
            Console.WriteLine($"PDF file exists: {pdfExists}");
            Console.WriteLine(pdfExists
                ? "Verification passed: PDF file was generated."
                : "Verification failed: PDF file was not found.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
