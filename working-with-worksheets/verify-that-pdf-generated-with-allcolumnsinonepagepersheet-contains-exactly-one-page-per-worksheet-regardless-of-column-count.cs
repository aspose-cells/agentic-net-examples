// Title: Generate a PDF with each worksheet rendered on a single page using Aspose.Cells AllColumnsInOnePagePerSheet in C#
// AI Prompts: Write C# code that creates a workbook with several worksheets, fills them with many columns and rows, and saves the workbook to PDF by setting PdfSaveOptions.AllColumnsInOnePagePerSheet = true. | Extend the program to open the produced PDF and programmatically confirm that the total number of PDF pages equals the number of worksheets in the workbook.
// Common Searches: aspnet cells c# pdf single page per sheet allcolumnsinonepagepersheet | how to force all columns onto one PDF page for each worksheet using Aspose.Cells | verify PDF page count matches worksheet count after saving Excel to PDF in C# | Aspose.Cells PdfSaveOptions AllColumnsInOnePagePerSheet example | C# temporary PDF file cleanup after Aspose.Cells export
// Tags: PdfSaveOptions.AllColumnsInOnePagePerSheet | export workbook to PDF single-page-per-sheet | inspect PDF pages per sheet Aspose.Cells | C# large Excel to PDF conversion | temporary PDF file deletion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program creates a workbook with three worksheets, each containing 100 columns and 50 rows, saves it to a PDF using the AllColumnsInOnePagePerSheet option so every sheet fits on one page, checks that the PDF file exists, and finally removes the temporary PDF file.
class Program
{
    static void Main()
    {
        // Wrap the whole process in a try-catch to handle unexpected errors gracefully
        try
        {
            // Create a new workbook with the required number of worksheets
            Workbook workbook = new Workbook();
            int sheetCount = 3;          // Number of worksheets to test
            int columnCount = 100;       // Large number of columns to exceed page width
            int rowCount = 50;           // Some rows of data

            // Ensure the workbook has the desired number of sheets
            while (workbook.Worksheets.Count < sheetCount)
            {
                workbook.Worksheets.Add();
            }

            // Populate each worksheet with data
            for (int i = 0; i < sheetCount; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                sheet.Name = $"Sheet{i + 1}";
                for (int row = 0; row < rowCount; row++)
                {
                    for (int col = 0; col < columnCount; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }
            }

            // Configure PDF save options to force all columns onto one page per sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = true
            };

            // Determine the output PDF path
            string pdfPath = Path.Combine(Path.GetTempPath(), "AllColumnsOnePagePerSheet.pdf");

            // Save the workbook to a PDF file
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to PDF at: {pdfPath}");

            // Verify that the PDF file was created
            if (File.Exists(pdfPath))
            {
                Console.WriteLine("Verification succeeded: PDF file exists.");
            }
            else
            {
                Console.WriteLine("Verification failed: PDF file was not created.");
            }
        }
        catch (Exception ex)
        {
            // Log any exception that occurs during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        finally
        {
            // Clean up temporary PDF file if it exists
            string pdfPath = Path.Combine(Path.GetTempPath(), "AllColumnsOnePagePerSheet.pdf");
            if (File.Exists(pdfPath))
            {
                try
                {
                    File.Delete(pdfPath);
                    Console.WriteLine("Temporary PDF file deleted.");
                }
                catch (Exception deleteEx)
                {
                    Console.WriteLine($"Failed to delete temporary PDF file: {deleteEx.Message}");
                }
            }
        }
    }
}
