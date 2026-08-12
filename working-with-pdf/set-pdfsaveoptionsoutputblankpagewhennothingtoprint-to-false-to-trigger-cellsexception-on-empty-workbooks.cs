// Title: Aspose.Cells C# – Disable Blank PDF Page and Trigger CellsException for Empty Workbooks
// Description: A C# sample that builds a workbook with no printable sheets, sets PdfSaveOptions.OutputBlankPageWhenNothingToPrint to false, and saves it as PDF. With no visible data, Aspose.Cells throws a CellsException, which the code captures, allowing developers to detect and block blank PDF output.
// Keywords: Aspose.Cells | PdfSaveOptions | OutputBlankPageWhenNothingToPrint | CellsException | empty workbook PDF | C# PDF export | prevent blank page | Aspose.Cells PDF exception | Excel to PDF conversion .NET | Aspose.Cells example GitHub
// Common Searches: Aspose.Cells throw exception when saving empty workbook to PDF | PdfSaveOptions OutputBlankPageWhenNothingToPrint false example | how to stop blank page generation in Aspose.Cells PDF export | catch CellsException for empty Excel workbook PDF conversion | C# Aspose.Cells PDF save options for empty worksheets
// Developer Intent: Configure PdfSaveOptions so that saving a workbook without printable content raises a CellsException instead of producing a blank PDF.
// Use Cases: Validate workbook content before PDF conversion and abort if nothing is printable. | Add error handling in automated reporting services to prevent empty PDF files. | Implement a safeguard in web APIs that convert Excel files to PDF, ensuring an exception is raised for fully hidden or empty sheets.
// AI Prompts: Write C# code that sets PdfSaveOptions.OutputBlankPageWhenNothingToPrint = false and catches the CellsException when an empty workbook is saved as PDF using Aspose.Cells. | Explain the effect of OutputBlankPageWhenNothingToPrint on PDF output in Aspose.Cells and describe best practices for handling the resulting CellsException. | Create unit tests in C# that verify a CellsException is thrown when OutputBlankPageWhenNothingToPrint is false and the workbook has no visible data.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // A C# sample that builds a workbook with no printable sheets, sets PdfSaveOptions.OutputBlankPageWhenNothingToPrint to false, and saves it as PDF. With no visible data, Aspose.Cells throws a CellsException, which the code captures, allowing developers to detect and block blank PDF output.
    class Program
    {
        static void Main()
        {
            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Add a second worksheet so we can hide the original one
            Worksheet hiddenSheet = workbook.Worksheets[0];
            Worksheet visibleSheet = workbook.Worksheets[workbook.Worksheets.Add()];

            // Hide the first worksheet to simulate a workbook with nothing to print
            hiddenSheet.IsVisible = false;

            // The second worksheet remains visible but contains no data

            // Configure PDF save options: do NOT output a blank page when nothing to print
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OutputBlankPageWhenNothingToPrint = false
            };

            try
            {
                // Attempt to save the workbook as PDF.
                // With OutputBlankPageWhenNothingToPrint set to false,
                // Aspose.Cells will not generate a blank page for an empty visible sheet.
                workbook.Save("EmptyWorkbook.pdf", pdfOptions);
                Console.WriteLine("PDF saved successfully.");
            }
            catch (CellsException ex)
            {
                // Handle any CellsException that may occur
                Console.WriteLine($"CellsException caught: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected exceptions
                Console.WriteLine($"Unexpected exception: {ex.Message}");
            }
        }
    }
}
