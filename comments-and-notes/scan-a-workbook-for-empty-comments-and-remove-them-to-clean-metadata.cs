// Title: Remove Empty Comments from an Excel Workbook Using Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx file, scans every worksheet for comments whose Note is null, empty, or whitespace, deletes those comments by cell address, and saves a cleaned copy. Ideal for reducing file size and eliminating placeholder notes.
// Keywords: Aspose.Cells | C# | remove empty comments | delete blank notes Excel | comment cleanup Aspose | Excel workbook preprocessing | remove placeholder comments | Aspose.Cells API comment removal
// Common Searches: how to delete empty comments in Excel using Aspose.Cells C# | remove blank notes from all worksheets Aspose.Cells | Aspose.Cells code to clean up comments with no text | C# script to purge empty Excel comments | Aspose.Cells remove comments without content
// Developer Intent: Identify and delete every comment that contains no visible text across all worksheets in an Excel workbook.
// Use Cases: Prepare a distribution‑ready workbook by stripping placeholder or accidental empty comments. | Automate preprocessing of user‑generated spreadsheets to keep file size minimal. | Standardize reporting pipelines by ensuring no blank notes remain before archiving or publishing.
// AI Prompts: Write C# code with Aspose.Cells that removes comments whose Note property is empty or whitespace on all worksheets. | Extend the sample to also delete comments that consist only of line‑break characters or invisible Unicode spaces. | Provide a step‑by‑step guide to log the cell addresses of removed comments before saving the workbook.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an .xlsx file, scans every worksheet for comments whose Note is null, empty, or whitespace, deletes those comments by cell address, and saves a cleaned copy. Ideal for reducing file size and eliminating placeholder notes.
class RemoveEmptyComments
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            CommentCollection comments = worksheet.Comments;
            List<string> emptyCommentCellNames = new List<string>();

            // Identify comments whose Note (text) is empty or whitespace
            foreach (Comment comment in comments)
            {
                if (string.IsNullOrWhiteSpace(comment.Note))
                {
                    // Convert row/column indices to cell name (e.g., "A1")
                    string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                    emptyCommentCellNames.Add(cellName);
                }
            }

            // Remove the empty comments using their cell names
            foreach (string cellName in emptyCommentCellNames)
            {
                comments.RemoveAt(cellName);
            }
        }

        // Save the cleaned workbook
        workbook.Save("output_cleaned.xlsx", SaveFormat.Xlsx);
    }
}
