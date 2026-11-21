// Title: Convert a TXT file to PDF with a page break after each paragraph using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert txt file to pdf with page break after each line | How to add a horizontal page break after each row when exporting to PDF with Aspose.Cells | Load plain text into an Aspose.Cells worksheet and save as PDF preserving paragraph breaks | Insert page breaks programmatically in Aspose.Cells before PDF conversion | Aspose.Cells txt to pdf conversion with paragraph spacing in C#
// Tags: txt to workbook loading Aspose.Cells | horizontal page break insertion Aspose.Cells | pdf export with page breaks Aspose.Cells | auto-detect text format Aspose.Cells | row based paragraph breaks Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the existence of an input.txt file, loads it into an Aspose.Cells Workbook (auto‑detecting the TXT format), adds a horizontal page break after each row to treat rows as paragraphs, and then saves the workbook as a PDF file using the SaveFormat.Pdf option.
class Program
{
    static void Main()
    {
        const string inputFile = "input.txt";
        const string outputFile = "output.pdf";

        try
        {
            // Verify that the input TXT file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the TXT file into a workbook (format is auto‑detected)
            Workbook workbook = new Workbook(inputFile);

            // Work with the first worksheet (the TXT data is placed here)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last row that contains data
            int lastRow = sheet.Cells.MaxDataRow;

            // Insert a horizontal page break after each row (treat each row as a paragraph)
            for (int row = 0; row <= lastRow; row++)
            {
                // Add a page break after the current row
                sheet.HorizontalPageBreaks.Add(row + 1);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputFile, SaveFormat.Pdf);
            Console.WriteLine($"PDF file saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
