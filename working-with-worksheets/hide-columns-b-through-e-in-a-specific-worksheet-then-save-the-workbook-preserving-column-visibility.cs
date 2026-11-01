// Title: Hide columns B through E in a worksheet and retain the hidden state when saving the workbook with Aspose.Cells for .NET
// AI Prompts: Hide a consecutive range of columns (B‑E) in a selected worksheet and save the workbook so the columns remain hidden using the Aspose.Cells API. | Iterate over column indexes to set their visibility to hidden and write the updated workbook to a new file with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells hide columns B to E in C# example | keep hidden columns after saving Excel workbook with Aspose.Cells | programmatically hide multiple adjacent columns using Aspose.Cells for .NET | worksheet.Cells.HideColumn range loop C# Aspose.Cells | preserve column visibility when exporting workbook Aspose.Cells
// Tags: hide column range Aspose.Cells | worksheet.Cells.HideColumn C# | preserve hidden columns on workbook save Aspose.Cells | Aspose.Cells column visibility Excel | hide columns B-E programmatically

using Aspose.Cells;
using System;

// // Loads input.xlsx, hides columns B through E (indexes 1‑4) in Sheet1 using worksheet.Cells.HideColumn, and saves the modified workbook to output.xlsx while preserving the hidden column state.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the target worksheet (change the name or index as needed)
        Worksheet worksheet = workbook.Worksheets["Sheet1"]; // or workbook.Worksheets[0]

        // Hide columns B (index 1) through E (index 4)
        for (int col = 1; col <= 4; col++)
        {
            worksheet.Cells.HideColumn(col);
        }

        // Save the workbook; column visibility is preserved in the saved file
        workbook.Save("output.xlsx");
    }
}
