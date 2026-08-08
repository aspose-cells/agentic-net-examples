// Title: Batch copy a template row to multiple rows across worksheets with Aspose.Cells for .NET
// Description: Loads a workbook, treats row 2 of the first worksheet as a template, and loops through the remaining sheets to copy that row to specified rows (e.g., 5, 10, 15) using Cells.CopyRow, preserving data and formatting, then saves the updated file.
// Keywords: Aspose.Cells | CopyRow | C# | .NET | batch copy rows | template row | copy row across worksheets | Excel automation | loop copy rows | preserve formatting
// Common Searches: How to copy a row to several rows in other worksheets using Aspose.Cells C# | Batch copy a template row across multiple sheets in .NET | Copy row with formatting to multiple destinations in Excel with Aspose.Cells | Loop through worksheets and duplicate a specific row in C# | Aspose.Cells copy row to many rows example
// Developer Intent: Copy a single template row from the first worksheet to a set of target rows in every other worksheet of the workbook.
// Use Cases: Apply a styled header row to all sheets in a report workbook. | Duplicate a pre‑formatted data‑entry row across several sections of each worksheet. | Insert a summary row template into every sheet for consistent layout.
// AI Prompts: Generate C# code that uses Aspose.Cells to copy row 2 from the first worksheet to rows 3, 7 and 12 in all other worksheets, keeping formatting intact. | Explain how to modify the loop to skip worksheets whose names contain the word "Archive" while copying the template row. | Show how to copy a row and then auto‑fit column widths in each target worksheet after the copy operation.

using System;
using Aspose.Cells;

// Loads a workbook, treats row 2 of the first worksheet as a template, and loops through the remaining sheets to copy that row to specified rows (e.g., 5, 10, 15) using Cells.CopyRow, preserving data and formatting, then saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook that contains the template row.
        // The template row is assumed to be in the first worksheet (index 0) at row index 1 (second row).
        Workbook workbook = new Workbook("Template.xlsx");

        // Source worksheet and its cells.
        Worksheet sourceSheet = workbook.Worksheets[0];
        Cells sourceCells = sourceSheet.Cells;
        int templateRowIndex = 1; // zero‑based index of the template row.

        // Define the destination row indices where the template row will be copied.
        int[] destinationRows = new int[] { 5, 10, 15 }; // example target rows.

        // Iterate over all worksheets except the source (template) sheet.
        for (int wsIndex = 1; wsIndex < workbook.Worksheets.Count; wsIndex++)
        {
            Worksheet targetSheet = workbook.Worksheets[wsIndex];
            Cells targetCells = targetSheet.Cells;

            // Copy the template row to each specified destination row in the current worksheet.
            foreach (int destRowIndex in destinationRows)
            {
                // CopyRow copies data and formats of the whole row from sourceCells to targetCells.
                targetCells.CopyRow(sourceCells, templateRowIndex, destRowIndex);
            }
        }

        // Save the workbook with the copied rows.
        workbook.Save("Result.xlsx");
    }
}
