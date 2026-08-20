// Title: C# – Convert TXT to PDF with a Page Break after Each Paragraph using Aspose.Cells
// Description: Load a plain‑text file as an Aspose.Cells workbook, add a horizontal page break after every row (paragraph), and export the sheet to PDF with PdfSaveOptions. The sample includes file‑existence checking and exception handling.
// Keywords: Aspose.Cells C# txt to pdf | add page break Aspose.Cells | horizontal page break worksheet | PdfSaveOptions Aspose.Cells | convert text file to PDF .NET | Excel page break programmatically | Aspose.Cells error handling
// Common Searches: Aspose.Cells convert txt file to PDF C# | how to insert page break after each row in Aspose.Cells | C# export text workbook to PDF with page breaks | Aspose.Cells PdfSaveOptions example | add horizontal page break before PDF export Aspose.Cells
// Developer Intent: Generate a PDF from a TXT workbook where each paragraph starts on a new page by inserting page breaks programmatically with Aspose.Cells.
// Use Cases: Create paginated PDF reports from plain‑text logs, with one line per page. | Automate document formatting where each paragraph must appear on a separate PDF page. | Validate input files and gracefully handle conversion errors in .NET applications.
// AI Prompts: Write C# code that reads a .txt file into an Aspose.Cells workbook, adds a page break after each row, and saves it as a PDF. | Explain the AddPageBreaks method in Aspose.Cells and how to use it before PDF conversion. | Provide best‑practice error handling for converting a text workbook to PDF with Aspose.Cells in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTxtToPdf
{
    // Load a plain‑text file as an Aspose.Cells workbook, add a horizontal page break after every row (paragraph), and export the sheet to PDF with PdfSaveOptions. The sample includes file‑existence checking and exception handling.
    public class TxtToPdfConverter
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.txt";
                const string outputPath = "output.pdf";

                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the TXT workbook; Aspose.Cells auto‑detects the format from the extension
                Workbook workbook = new Workbook(inputPath);

                // Work with the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the last row that contains data (zero‑based index)
                int lastDataRow = sheet.Cells.MaxDataRow;

                // Insert a horizontal page break after each data row
                for (int row = 0; row <= lastDataRow; row++)
                {
                    // Break after the current row: Excel rows are 1‑based
                    int breakRow = row + 2; // move to the first cell of the next row
                    string cellName = $"A{breakRow}";
                    sheet.AddPageBreaks(cellName);
                }

                // Save the workbook as PDF
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main()
        {
            Run();
        }
    }
}
