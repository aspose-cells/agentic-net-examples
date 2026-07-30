// Title: C# – Hide Rows in an Aspose.Cells Workbook by Cell Value and Save as XLSX
// Description: Creates a workbook, populates column A with sample strings, hides every row whose cell equals "Hide", verifies hidden status, and saves the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells hide rows C# | conditional row hiding Aspose.Cells | save workbook after hiding rows | filter rows by cell value .NET | Aspose.Cells row visibility API
// Common Searches: how to hide rows in Aspose.Cells based on cell content | C# hide worksheet rows when column value matches | Aspose.Cells save workbook after hiding rows | check if a row is hidden Aspose.Cells C#
// Developer Intent: Programmatically conceal rows that meet a specific cell‑value condition and persist the modified workbook.
// Use Cases: Prepare a financial report that automatically hides rows flagged as "Hide" before distribution. | Export data sets while suppressing rows marked with a status flag, keeping the original sheet intact. | Run a validation step that confirms hidden rows before finalizing the XLSX output.
// AI Prompts: Generate C# code with Aspose.Cells to hide rows where column B contains "Inactive" and then save the file. | Explain how to toggle row hidden state in Aspose.Cells and retain it when reopening the workbook. | Show an example of using a custom predicate function to hide rows in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a workbook, populates column A with sample strings, hides every row whose cell equals "Hide", verifies hidden status, and saves the result as an XLSX file using Aspose.Cells for .NET.
class HideRowsBasedOnPredicate
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data in column A
        string[] values = { "Keep", "Hide", "Keep", "Hide", "Keep" };
        for (int i = 0; i < values.Length; i++)
        {
            cells[i, 0].PutValue(values[i]); // Row i, Column A (0‑based index)
        }

        // Hide rows where the cell value equals "Hide"
        for (int row = 0; row < values.Length; row++)
        {
            if (cells[row, 0].StringValue == "Hide")
            {
                cells.HideRow(row);
            }
        }

        // Optional: display hidden status for verification
        for (int row = 0; row < values.Length; row++)
        {
            Console.WriteLine($"Row {row + 1} hidden: {cells.IsRowHidden(row)}");
        }

        // Save the modified workbook
        workbook.Save("HiddenRowsExample.xlsx", SaveFormat.Xlsx);
    }
}
