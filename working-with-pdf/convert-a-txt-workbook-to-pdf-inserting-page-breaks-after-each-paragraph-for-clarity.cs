// Title: Convert a TXT file to PDF with a page break after each line using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a .txt file, writes each line into separate rows of an Aspose.Cells worksheet, inserts a horizontal page break after each row, and saves the workbook as a PDF. | Show how to add page breaks between rows when exporting a text‑based worksheet to PDF with Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells C# add horizontal page break after each row when saving workbook as PDF | How to convert a plain text file to PDF with line‑by‑line page breaks using Aspose.Cells | C# example for reading a .txt file into a worksheet and inserting page breaks before PDF export | Save text workbook as PDF with page breaks per paragraph Aspose.Cells .NET
// Tags: Aspose.Cells insert horizontal page breaks | C# convert txt to PDF with Aspose.Cells | Aspose.Cells workbook to PDF with page breaks | read text file into worksheet Aspose.Cells | save worksheet as PDF Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program reads a .txt file, writes each line into successive rows of a new Aspose.Cells workbook, adds a horizontal page break after every populated row, ensures the output folder exists, and then saves the workbook as a PDF document.
class Program
{
    static void Main()
    {
        const string inputPath = "input.txt";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Read all lines from the text file; each line will become a row in the worksheet
            string[] lines = File.ReadAllLines(inputPath);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with the text lines
            for (int i = 0; i < lines.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(lines[i]);
            }

            // Insert a horizontal page break after each populated row
            // (row index is zero‑based, so we add a break at row+1)
            for (int row = 0; row < lines.Length; row++)
            {
                sheet.HorizontalPageBreaks.Add(row + 1);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully created at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
