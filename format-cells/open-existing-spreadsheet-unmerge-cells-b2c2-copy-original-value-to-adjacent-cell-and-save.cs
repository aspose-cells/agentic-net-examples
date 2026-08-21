// Title: C# – Unmerge B2:C2, copy merged value to D2, and save workbook with Aspose.Cells
// Description: Load an existing Excel file, retrieve the value from the merged range B2:C2, unmerge the cells, copy the original value to D2, and save the updated workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | unmerge cells | merged cell value | copy cell value | Excel range B2:C2 | write to D2 | Workbook.Save | Excel automation | Aspose.Cells API
// Common Searches: Aspose.Cells unmerge specific range and keep value | C# copy value from merged cell after unmerge | How to move merged cell content to another cell with Aspose.Cells | Unmerge B2:C2 and write value to D2 in .NET | Save edited Excel workbook using Aspose.Cells
// Developer Intent: Extract the original value from B2:C2, unmerge the range, place that value into D2, and persist the changes.
// Use Cases: Cleaning imported spreadsheets that contain merged headers before data import. | Preparing Excel files for systems that reject merged cells while preserving header text. | Generating a version of a report where each column has its own cell with the original merged content.
// AI Prompts: Generate C# code with Aspose.Cells to unmerge a given range and copy its original value to another cell. | Explain how to preserve merged‑cell values when converting an Aspose.Cells workbook to CSV. | Provide a step‑by‑step tutorial for unmerging cells, duplicating their content, and saving the workbook under a new name using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Load an existing Excel file, retrieve the value from the merged range B2:C2, unmerge the cells, copy the original value to D2, and save the updated workbook using Aspose.Cells for .NET.
class UnmergeAndCopy
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Get the value from the merged cell (top‑left cell of the range B2:C2)
        Cell mergedCell = cells["B2"];
        object originalValue = mergedCell.Value;

        // Unmerge the range B2:C2
        // B2 is row 1, column 1 (zero‑based). The range spans 1 row and 2 columns.
        cells.UnMerge(1, 1, 1, 2);

        // Copy the original value to the adjacent cell D2 (row 1, column 3)
        cells["D2"].PutValue(originalValue);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
