// Title: Add Excel cell comments as footnote entries on PDF pages using Aspose.Cells for .NET
// AI Prompts: Create a C# console program that loads an .xlsx file, extracts every worksheet comment, writes each comment with its sheet and cell reference into a new 'Footnotes' worksheet, and then saves the workbook as a PDF with Aspose.Cells. | Write a C# method that iterates through all worksheets in a Workbook, gathers comment text and addresses, populates column A of a newly added worksheet named 'Footnotes', and uses PdfSaveOptions to generate a PDF file. | Develop a C# script that verifies the presence of an input Excel file, formats each comment as 'SheetName!Cell: Comment', adds these entries to a footnote sheet, and outputs a PDF document using Aspose.Cells.
// Common Searches: how to include Excel comments as footnotes when converting to PDF with Aspose.Cells C# | Aspose.Cells C# extract cell comments and add them to a footnote worksheet before PDF export | convert workbook to PDF and list all comments in a separate sheet using Aspose.Cells .NET | C# generate PDF from Excel with comment references on each page using Aspose.Cells | Aspose.Cells save Excel as PDF with a footnotes sheet containing comment details
// Tags: extract Excel cell comments with Aspose.Cells | add footnotes worksheet for PDF conversion | PdfSaveOptions comment footnotes Aspose.Cells | C# generate PDF from workbook with comment references | Aspose.Cells create footnotes from comments

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// // Loads an Excel file, gathers all cell comments, writes them as 'Sheet!Cell: Comment' entries in a new 'Footnotes' worksheet, and saves the workbook as a PDF using Aspose.Cells.
class ExcelCommentsToPdfFootnotes
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Collect all comments with their cell addresses
            List<string> footnotes = new List<string>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Comment comment in sheet.Comments)
                {
                    // Determine the cell address of the comment
                    int row = comment.Row;
                    int col = comment.Column;
                    string cellName = CellsHelper.CellIndexToName(row, col);

                    // Build a footnote entry: SheetName!CellAddress: CommentText
                    string entry = $"{sheet.Name}!{cellName}: {comment.Note}";
                    footnotes.Add(entry);
                }
            }

            // If there are comments, add a new worksheet to hold the footnotes
            if (footnotes.Count > 0)
            {
                int footnoteSheetIndex = workbook.Worksheets.Add();
                Worksheet footnoteSheet = workbook.Worksheets[footnoteSheetIndex];
                footnoteSheet.Name = "Footnotes";

                // Write each footnote into column A
                for (int i = 0; i < footnotes.Count; i++)
                {
                    footnoteSheet.Cells[i, 0].PutValue(footnotes[i]);
                }
            }

            // Save the workbook as PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
