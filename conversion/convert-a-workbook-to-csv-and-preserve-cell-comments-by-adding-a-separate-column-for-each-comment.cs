// Title: Convert Excel to CSV with each comment in its own column using Aspose.Cells for .NET
// Description: Loads an XLSX workbook, scans every worksheet for cell comments, inserts a new column next to each commented cell, writes the comment text into that column, optionally removes the original comment, and saves the result as a single CSV file that includes all sheets.
// Keywords: Aspose.Cells CSV export with comments | C# preserve Excel comments in CSV | insert column for comment Aspose.Cells | convert XLSX to CSV retaining notes | export all worksheets to one CSV | Aspose.Cells comment extraction | Excel to CSV with annotations | .NET workbook to CSV
// Common Searches: how to keep Excel comments when converting to CSV using Aspose.Cells | C# add column for each cell comment before CSV export | Aspose.Cells save workbook as CSV with comment text | export multiple sheets to single CSV preserving notes | insert comment columns in Excel with Aspose.Cells
// Developer Intent: The developer needs to transform an Excel workbook into a CSV file while capturing every cell comment as a separate column in the output.
// Use Cases: Generating CSV audit logs that include user remarks originally stored as Excel comments. | Feeding downstream systems with comment data as distinct fields for data enrichment. | Archiving Excel worksheets as CSV for compliance, ensuring annotation visibility.
// AI Prompts: Create C# code with Aspose.Cells that converts a workbook to CSV and adds a new column for each cell comment, preserving row order. | Show how to modify the sample so comment columns are appended at the end of each sheet instead of after the original cells. | Provide a method to handle rows that contain several comments when exporting to CSV, ensuring no column name collisions.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells;

// Loads an XLSX workbook, scans every worksheet for cell comments, inserts a new column next to each commented cell, writes the comment text into that column, optionally removes the original comment, and saves the result as a single CSV file that includes all sheets.
class WorkbookToCsvWithComments
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the collection of comments on the current worksheet
            CommentCollection comments = sheet.Comments;

            // If there are no comments, continue to the next sheet
            if (comments.Count == 0) continue;

            // Prepare a list of comment information (row, column, note)
            var commentInfo = new List<(int Row, int Column, string Note)>();
            foreach (Comment comment in comments)
            {
                commentInfo.Add((comment.Row, comment.Column, comment.Note));
            }

            // Sort comments by column index descending to avoid shifting issues when inserting columns
            var sortedComments = commentInfo.OrderByDescending(c => c.Column).ToList();

            // Insert a new column for each comment and place the comment text in that column
            foreach (var (row, column, note) in sortedComments)
            {
                // Insert a new column right after the original column
                sheet.Cells.InsertColumn(column + 1);

                // Write the comment text into the newly inserted column, same row as the original cell
                sheet.Cells[row, column + 1].PutValue(note);
            }

            // Optional: clear original comments if they are no longer needed
            // sheet.ClearComments();
        }

        // Prepare CSV save options to export all worksheets into a single CSV file
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            ExportAllSheets = true
        };

        // Save the modified workbook as CSV
        workbook.Save("output.csv", csvOptions);
    }
}
