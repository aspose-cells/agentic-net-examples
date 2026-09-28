// Title: Load an Excel workbook, set the first worksheet’s paper size to Letter, export to PDF in a MemoryStream, and confirm the PDF stream is not empty using Aspose.Cells for .NET (C#)
// AI Prompts: Load a .xlsx file with Aspose.Cells, change the first worksheet's PageSetup.PaperSize to PaperLetter, save the workbook as PDF into a MemoryStream, and output the stream length. | Write a C# snippet that opens an Excel file, applies a Letter paper size to the first sheet, converts the workbook to PDF in memory, and asserts that the resulting PDF contains exactly one page. | Create a unit test that uses Aspose.Cells to load a workbook, set the worksheet paper size to Letter via PageSetup, saves to PDF with PdfSaveOptions, and verifies the PDF stream size is greater than zero.
// Common Searches: Aspose.Cells C# set worksheet paper size to Letter before PDF conversion | How to export Excel to PDF in a MemoryStream using Aspose.Cells .NET | Validate PDF page count after converting an Excel workbook with Aspose.Cells | Check PDF stream length after saving workbook as PDF in C# Aspose.Cells
// Tags: set worksheet paper size Letter Aspose.Cells | export workbook to PDF memory stream C# | verify PDF stream length Aspose.Cells | convert Excel to PDF Aspose.Cells .NET | page count validation PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, sets the first worksheet's paper size to Letter via PageSetup, saves the workbook as a PDF into a MemoryStream using PdfSaveOptions, and confirms that the generated PDF stream contains data, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the paper size for the first worksheet (Letter)
            if (workbook.Worksheets.Count > 0)
            {
                workbook.Worksheets[0].PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }

            // Set PDF save options (no PaperSize property here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save to PDF in a memory stream
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, pdfOptions);
                pdfStream.Position = 0; // Reset for further use

                // Verify that the PDF stream contains data
                if (pdfStream.Length > 0)
                {
                    Console.WriteLine($"PDF generation succeeded. Stream length: {pdfStream.Length}");
                }
                else
                {
                    Console.WriteLine("PDF generation failed: resulting stream is empty.");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
