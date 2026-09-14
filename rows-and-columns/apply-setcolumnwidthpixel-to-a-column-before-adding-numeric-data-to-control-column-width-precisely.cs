// Title: How to set a worksheet column width in pixels before writing numeric values using Aspose.Cells for .NET (C#)
// AI Prompts: Set column B to 150 pixels with cells.SetColumnWidthPixel, then fill rows 0‑4 with incremental numbers and save the workbook. | Create a new Workbook, adjust the pixel width of a specific column, populate it with calculated numeric data, and export to an .xlsx file using Aspose.Cells C#. | Show how to apply SetColumnWidthPixel to define exact pixel width, then write numeric entries and persist the file.
// Common Searches: Aspose.Cells C# set column width in pixels before adding data to worksheet | SetColumnWidthPixel example for numeric data insertion in .NET | How to control column width in pixels with Aspose.Cells and then write numbers | C# Aspose.Cells set column B width 150px then populate values | Precise column sizing using SetColumnWidthPixel in Aspose.Cells
// Tags: SetColumnWidthPixel pixel column sizing | Aspose.Cells column width pre‑population | C# define column width in pixels | populate numeric values after column sizing | save workbook .xlsx Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, sets column B's width to 150 pixels using SetColumnWidthPixel, fills rows 0‑4 of that column with multiples of 10, and saves the file as ColumnWidthPixelDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Set the width of column B (index 1) to 150 pixels
        cells.SetColumnWidthPixel(1, 150);

        // Populate column B with numeric data
        for (int row = 0; row < 5; row++)
        {
            cells[row, 1].PutValue(row * 10); // Example values: 0,10,20,30,40
        }

        // Save the workbook
        workbook.Save("ColumnWidthPixelDemo.xlsx");
    }
}
