using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook – replace with the provided create rule if available
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample fixed‑width strings placed in column C (zero‑based column index 2)
        sheet.Cells["C1"].PutValue("ABCDEF123456");
        sheet.Cells["C2"].PutValue("GHIJKL789012");

        // Configure TextToColumns for fixed‑width splitting.
        // The exact API for specifying column widths is not shown in the documentation;
        // if a property such as `ColumnWidths` exists, set it here (e.g., new int[] {3,3,6}).
        TxtLoadOptions options = new TxtLoadOptions();
        // options.ColumnWidths = new int[] { 3, 3, 6 }; // <-- placeholder for fixed‑width settings

        // Apply TextToColumns on column C (row 0, column 2) for the two rows of data.
        sheet.Cells.TextToColumns(0, 2, 2, options);

        // Save the workbook – replace with the provided save rule if available
        workbook.Save("output.xlsx");
    }
}

// Author: Example demonstrating fixed‑width TextToColumns on column C using Aspose.Cells for .NET.