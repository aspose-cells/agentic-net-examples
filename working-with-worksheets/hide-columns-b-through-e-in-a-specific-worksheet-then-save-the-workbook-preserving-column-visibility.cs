// Title: C# – Hide Columns B‑E in an Aspose.Cells Worksheet and Save with Hidden State
// Description: Creates a new Workbook, accesses the first Worksheet, optionally fills sample data, hides columns B through E using worksheet.Cells.HideColumns(1, 4) (zero‑based indices), and saves the file as HiddenColumns.xlsx while preserving the hidden column visibility.
// Keywords: Aspose.Cells hide columns C# | HideColumns method | Excel column visibility | C# Aspose.Cells example | Save workbook with hidden columns
// Common Searches: how to hide columns B to E using Aspose.Cells .NET | Aspose.Cells hide multiple columns by index | save Excel file with hidden columns Aspose | C# hide worksheet columns programmatically
// Developer Intent: Hide columns B‑E in a worksheet and persist the hidden state when saving the workbook.
// Use Cases: Prepare a report that omits sensitive data columns before distribution. | Create a template that shows only relevant columns to end‑users while keeping the rest hidden. | Improve readability of large datasets by collapsing auxiliary columns prior to export.
// AI Prompts: Generate C# code with Aspose.Cells to hide columns C‑G in the second worksheet and save as .xlsb, keeping the columns hidden. | Explain the zero‑based indexing of HideColumns in Aspose.Cells and demonstrate how to unhide columns later in C#.

using System;
using Aspose.Cells;

// Creates a new Workbook, accesses the first Worksheet, optionally fills sample data, hides columns B through E using worksheet.Cells.HideColumns(1, 4) (zero‑based indices), and saves the file as HiddenColumns.xlsx while preserving the hidden column visibility.
class HideColumnsExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (or specify by name/index)
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (optional, just for illustration)
        for (int row = 0; row < 10; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Hide columns B through E (zero‑based indices 1 to 4, total 4 columns)
        worksheet.Cells.HideColumns(1, 4);

        // Save the workbook while preserving column visibility
        workbook.Save("HiddenColumns.xlsx", SaveFormat.Xlsx);
    }
}
