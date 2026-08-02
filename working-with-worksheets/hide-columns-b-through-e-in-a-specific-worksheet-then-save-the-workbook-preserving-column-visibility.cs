// Title: Hide Columns B‑E in an Excel Worksheet with Aspose.Cells for .NET and Preserve Visibility on Save
// Description: Creates a new Workbook, populates sample data, hides columns B through E on the first worksheet using the HideColumns method, and saves the file as an .xlsx where the hidden column state is retained.
// Keywords: Aspose.Cells | C# | .NET | HideColumns | hide columns | Excel column visibility | preserve hidden columns | worksheet column hide | Excel automation | Aspose.Cells API
// Common Searches: Aspose.Cells hide columns B to E | C# hide multiple columns in Excel using Aspose | Save Excel file with hidden columns Aspose.Cells | How to hide a column range in Aspose.Cells .NET | Preserve column visibility when saving workbook Aspose
// Developer Intent: Hide columns B‑E in a specific worksheet and save the workbook while keeping those columns hidden.
// Use Cases: Generate a report that hides confidential or intermediate calculation columns before distribution. | Create an Excel template that automatically conceals helper columns for end‑users. | Export data to Excel from an application while ensuring certain columns remain invisible in the final file.
// AI Prompts: Write C# code with Aspose.Cells to hide columns C‑G and save the workbook as .xlsb. | Show how to unhide columns that were previously hidden in a worksheet using Aspose.Cells for .NET. | Provide an example that hides rows and columns based on a condition with Aspose.Cells. | Demonstrate how to hide columns by index range and then toggle their visibility programmatically.

using System;
using Aspose.Cells;

// Creates a new Workbook, populates sample data, hides columns B through E on the first worksheet using the HideColumns method, and saves the file as an .xlsx where the hidden column state is retained.
class HideColumnsExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the target worksheet (first worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data (optional, just for demonstration)
        for (int row = 0; row < 10; row++)
        {
            for (int col = 0; col < 6; col++)
            {
                worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Hide columns B through E.
        // Column indices are zero‑based: B = 1, C = 2, D = 3, E = 4.
        // HideColumns(startColumn, totalColumns) hides a range.
        worksheet.Cells.HideColumns(1, 4); // hides columns 1,2,3,4 (B‑E)

        // Save the workbook; hidden column states are preserved.
        workbook.Save("HiddenColumns_BtoE.xlsx", SaveFormat.Xlsx);
    }
}
