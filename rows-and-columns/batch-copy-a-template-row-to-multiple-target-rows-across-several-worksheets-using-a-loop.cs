// Title: Batch copy a template row to multiple rows across all worksheets with Aspose.Cells for .NET
// Description: Loads a workbook, creates worksheet copies, then loops through every sheet and uses Cells.CopyRows to duplicate a template row (row 0) into specified target rows (e.g., 2, 5, 8). The updated workbook is saved as a new file.
// Keywords: Aspose.Cells | C# copy row | Cells.CopyRows example | batch row copy | copy template row across worksheets | .NET spreadsheet automation | loop through worksheets | duplicate row multiple targets
// Common Searches: Aspose.Cells copy a row to many rows in all sheets | How to duplicate a template row across worksheets using C# | Batch copy rows with Aspose.Cells loop | CopyRows method multiple target rows example | Copy a header row to every worksheet in a workbook
// Developer Intent: Copy a single template row to several target rows in every worksheet of a workbook using Aspose.Cells for .NET.
// Use Cases: Add a consistent header row to all sheets in a reporting workbook. | Insert a pre‑formatted data‑validation row at multiple positions in each worksheet. | Replicate a formula row across several rows on every sheet for uniform calculations.
// AI Prompts: Generate C# code with Aspose.Cells that copies row 0 from the first worksheet to rows 3, 6, and 9 in all worksheets while preserving styles and formulas. | Explain how to modify the nested loop to copy multiple source rows to corresponding target rows across every worksheet. | Show how to read target row indices from a JSON file and use Aspose.Cells to batch copy the template row to those rows in each worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsBatchCopyRow
{
    // Loads a workbook, creates worksheet copies, then loops through every sheet and uses Cells.CopyRows to duplicate a template row (row 0) into specified target rows (e.g., 2, 5, 8). The updated workbook is saved as a new file.
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains the template row (row 0) in the first worksheet
            Workbook workbook = new Workbook("Template.xlsx");

            // Add a few copies of the template worksheet to demonstrate copying across multiple sheets
            // The AddCopy method copies the worksheet within the same workbook
            workbook.Worksheets.AddCopy("Sheet1"); // copy 1
            workbook.Worksheets.AddCopy("Sheet1"); // copy 2

            // Reference to the worksheet that holds the source (template) row
            Worksheet templateSheet = workbook.Worksheets[0];
            Cells templateCells = templateSheet.Cells;
            int templateRowIndex = 0; // zero‑based index of the template row

            // Define the target row indices where the template row should be copied
            int[] targetRows = new int[] { 2, 5, 8 }; // example rows (zero‑based)

            // Loop through all worksheets in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells destCells = ws.Cells;

                // Copy the template row to each target row within the current worksheet
                foreach (int targetRow in targetRows)
                {
                    // CopyRows(sourceCells, sourceRowIndex, destinationRowIndex, rowNumber)
                    destCells.CopyRows(templateCells, templateRowIndex, targetRow, 1);
                }
            }

            // Save the modified workbook
            workbook.Save("BatchCopyResult.xlsx");
        }
    }
}
