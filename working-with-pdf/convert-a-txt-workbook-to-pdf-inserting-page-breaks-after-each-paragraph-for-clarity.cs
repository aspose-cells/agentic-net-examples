// Title: Convert a TXT file to PDF with paragraph‑level page breaks using Aspose.Cells in C#
// AI Prompts: Write C# code that reads a .txt file into an Aspose.Cells workbook with TxtLoadOptions, identifies paragraphs separated by empty rows, inserts a horizontal page break after each paragraph, and saves the workbook as a PDF. | Show how to programmatically add horizontal page breaks in an Aspose.Cells worksheet based on consecutive non‑empty rows that form a paragraph. | Demonstrate robust error handling for missing input files and unexpected exceptions when converting a text workbook to PDF with Aspose.Cells.
// Common Searches: asp.net how to add page breaks after each paragraph when converting a text file to PDF with Aspose.Cells | c# detect paragraph boundaries in a worksheet loaded from a .txt file using Aspose.Cells | save workbook as PDF with horizontal page breaks based on empty rows Aspose.Cells | convert txt to pdf preserving paragraph spacing Aspose.Cells C# example
// Tags: Aspose.Cells TxtLoadOptions paragraph detection | C# insert horizontal page breaks Aspose.Cells | convert text workbook to PDF Aspose.Cells | page break insertion based on empty rows | error handling file not found Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for the existence of input.txt, loads it into a Workbook using TxtLoadOptions (each line becomes a row), scans rows to locate paragraph blocks separated by blank rows, adds a horizontal page break after each block, and saves the workbook as output.pdf while handling possible errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.txt";
            const string outputPath = "output.pdf";

            // Verify that the input TXT file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the TXT file into a workbook. Each line becomes a separate row.
            var loadOptions = new TxtLoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Work with the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last row that contains data.
            int lastDataRow = sheet.Cells.MaxDataRow;

            // Insert a horizontal page break after each paragraph.
            // A paragraph is defined as a consecutive block of non‑empty rows.
            int currentRow = 0;
            while (currentRow <= lastDataRow)
            {
                // Skip any leading empty rows.
                if (string.IsNullOrWhiteSpace(sheet.Cells[currentRow, 0].StringValue))
                {
                    currentRow++;
                    continue;
                }

                // Find the last row of the current paragraph.
                int paragraphEndRow = currentRow;
                while (paragraphEndRow + 1 <= lastDataRow &&
                       !string.IsNullOrWhiteSpace(sheet.Cells[paragraphEndRow + 1, 0].StringValue))
                {
                    paragraphEndRow++;
                }

                // Add a page break before the row that follows the paragraph.
                // HorizontalPageBreaks.Add inserts a break before the specified row index.
                sheet.HorizontalPageBreaks.Add(paragraphEndRow + 1);

                // Move to the row after the paragraph (which may be empty).
                currentRow = paragraphEndRow + 1;
            }

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
