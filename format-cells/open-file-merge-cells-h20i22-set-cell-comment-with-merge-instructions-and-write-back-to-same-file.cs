// Title: C# – Merge cells H20:I22 and add a comment using Aspose.Cells
// Description: Loads or creates Workbook.xlsx, merges the range H20:I22, inserts a comment on H20 describing the merge, and saves the workbook back to the same file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge cells C# | add comment merged range Aspose | update existing Excel file .NET | merge H20 I22 Aspose.Cells | Excel automation C# comment
// Common Searches: Aspose.Cells merge specific range and add comment C# | C# code to merge H20 to I22 in Excel | how to add a comment to a merged cell with Aspose.Cells | update existing workbook merge cells Aspose .NET
// Developer Intent: Combine a defined cell block in an existing Excel workbook and attach a descriptive comment using Aspose.Cells for .NET.
// Use Cases: Create a multi‑column header spanning rows 20‑22 and document its purpose. | Automate layout adjustments in a template workbook while providing annotation for downstream users. | Refresh a report by merging cells for visual grouping and adding an author‑attributed note.
// AI Prompts: Write C# code that opens Workbook.xlsx, merges cells H20:I22, adds a comment to H20, and saves the file with Aspose.Cells. | Provide a robust Aspose.Cells snippet that checks for workbook existence, creates it if missing, merges a range, adds a comment with author, and includes error handling.

using System;
using System.IO;
using Aspose.Cells;

// Loads or creates Workbook.xlsx, merges the range H20:I22, inserts a comment on H20 describing the merge, and saves the workbook back to the same file with Aspose.Cells for .NET.
class MergeCellsWithComment
{
    static void Main()
    {
        string filePath = "Workbook.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook or create a new one if the file is missing
            if (File.Exists(filePath))
            {
                workbook = new Workbook(filePath);
            }
            else
            {
                workbook = new Workbook(); // creates a workbook with a default worksheet
                workbook.Save(filePath);   // optional: persist the empty workbook for future runs
            }

            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells H20:I22 (zero‑based indices: rows 19‑21, columns 7‑8)
            cells.Merge(firstRow: 19, firstColumn: 7, totalRows: 3, totalColumns: 2);

            // Add a comment to the upper‑left cell of the merged range (H20)
            int commentIndex = worksheet.Comments.Add("H20");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "Cells H20:I22 have been merged.";
            comment.Author = "Automation";

            // Save the workbook with the changes
            workbook.Save(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
