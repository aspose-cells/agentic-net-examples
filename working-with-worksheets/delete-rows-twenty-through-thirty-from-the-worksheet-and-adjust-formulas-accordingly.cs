// Title: Remove rows 20‑30 from an Excel worksheet and let Aspose.Cells automatically adjust formulas in C#
// Common Searches: C# Aspose.Cells delete rows 20 to 30 and keep formulas correct | how to remove a range of rows in Excel using Aspose.Cells .NET | Aspose.Cells delete multiple rows and update dependent formulas in C# | example of Worksheet.Cells.DeleteRows for rows 20‑30 in Aspose.Cells | remove rows 20‑30 from first worksheet and recalculate formulas with Aspose.Cells
// Tags: Worksheet.Cells.DeleteRows API | remove rows range Aspose.Cells C# | formula auto-adjustment on row deletion | Excel worksheet bulk row removal Aspose.Cells | C# .NET Excel row deletion with formula update

using System;
using Aspose.Cells;

// Loads 'input.xlsx', deletes rows 20 through 30 (zero‑based index 19, 11 rows) from the first worksheet using Worksheet.Cells.DeleteRows, which automatically adjusts any formulas that reference the removed rows, and saves the modified workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete rows 20 through 30 (zero‑based index 19, total 11 rows)
        // Aspose.Cells automatically updates formulas that reference the deleted rows.
        worksheet.Cells.DeleteRows(19, 11);

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
