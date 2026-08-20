// Title: C# – Hide rows in an Aspose.Cells workbook based on cell values and save as XLSX
// Description: Creates a new workbook, fills column A with sample strings, hides every row whose A‑cell equals a specified value (e.g., "Hide"), and saves the modified file. Demonstrates using Cells.HideRow with a simple predicate in Aspose.Cells for .NET.
// Keywords: Aspose.Cells hide rows C# | hide Excel rows programmatically | Cells.HideRow example | filter rows by cell value Aspose | save workbook after hiding rows | C# Excel row visibility | Aspose.Cells conditional row hide
// Common Searches: how to hide rows in Aspose.Cells when a cell contains specific text | C# code to hide rows based on column A value using Aspose.Cells | Aspose.Cells hide rows and save workbook example | programmatically hide Excel rows with Aspose.Cells .NET | hide rows predicate Aspose.Cells C#
// Developer Intent: Programmatically conceal rows that meet a condition and persist the workbook.
// Use Cases: Generate reports that automatically hide rows flagged as "Hide" before distribution. | Create a clean view of data by removing rows with a particular status without deleting them. | Build Excel templates that react to a flag column, hiding rows dynamically during runtime.
// AI Prompts: Write C# code using Aspose.Cells to hide rows where column B contains the word "Inactive" and then save the workbook as XLSX. | Show an example that iterates through a worksheet and hides rows based on a custom predicate function with Aspose.Cells. | Provide a reusable method in C# that accepts a predicate and hides matching rows in an Aspose.Cells worksheet, then saves the file.

using System;
using Aspose.Cells;

// Creates a new workbook, fills column A with sample strings, hides every row whose A‑cell equals a specified value (e.g., "Hide"), and saves the modified file. Demonstrates using Cells.HideRow with a simple predicate in Aspose.Cells for .NET.
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
            // Populate cells A1, A2, ... (zero‑based index)
            cells[i, 0].PutValue(values[i]);
        }

        // Hide rows where the cell value equals "Hide"
        for (int i = 0; i < values.Length; i++)
        {
            if (cells[i, 0].StringValue == "Hide")
            {
                // Hide the row (zero‑based index)
                cells.HideRow(i);
            }
        }

        // Save the workbook with hidden rows applied
        workbook.Save("HiddenRowsExample.xlsx", SaveFormat.Xlsx);
    }
}
