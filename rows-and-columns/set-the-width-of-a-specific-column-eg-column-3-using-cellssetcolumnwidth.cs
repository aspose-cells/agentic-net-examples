using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (use the provided creation rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the width of column 3 (zero‑based index 2) to 25.5 characters
        worksheet.Cells.SetColumnWidth(2, 25.5);

        // Save the workbook (use the provided save rule)
        workbook.Save("ColumnWidthDemo.xlsx", SaveFormat.Xlsx);
    }
}

// Author: Example demonstrating Cells.SetColumnWidth to adjust column width.