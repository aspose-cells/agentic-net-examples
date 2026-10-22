// Title: How to add a SUBTOTAL formula and freeze the top rows in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that populates column B, inserts a SUBTOTAL(9,B2:B10) formula in B11, adds a "Subtotal" label in A11, freezes the first 11 rows, and saves the workbook as SubtotalsAndFreeze.xlsx. | Generate an Aspose.Cells example that demonstrates inserting a SUBTOTAL function that ignores hidden rows, then applies FreezePanes to keep the subtotal row visible while scrolling.
// Common Searches: Aspose.Cells C# insert SUBTOTAL function and freeze header rows | C# freeze panes after adding subtotal formula with Aspose.Cells | How to keep subtotal row visible in Excel using Aspose.Cells FreezePanes | Using Aspose.Cells to calculate subtotal that ignores hidden rows in .NET | Example of SUBTOTAL(9,…) with FreezePanes in Aspose.Cells for .NET
// Tags: Aspose.Cells insert SUBTOTAL formula | Aspose.Cells freeze panes rows | C# calculate subtotal ignoring hidden rows | Aspose.Cells generate Excel with frozen rows | SUBTOTAL function with FreezePanes in .NET

using Aspose.Cells;
using System;

// The program creates a new workbook, fills cells B2:B10 with incremental values, places a SUBTOTAL(9,B2:B10) formula in B11 with a "Subtotal" label in A11, freezes the first 11 rows to keep the subtotal visible, and saves the file as SubtotalsAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column B (cells B2:B10)
            for (int i = 1; i <= 9; i++) // rows 2‑10 (zero‑based index)
            {
                sheet.Cells[i, 1].PutValue(i * 10); // example values: 10, 20, ...
            }

            // Insert a SUBTOTAL formula in row 11 (cell B11)
            // Function_num 9 = SUM, which ignores hidden rows
            sheet.Cells[10, 1].Formula = "SUBTOTAL(9,B2:B10)";

            // Optional label for the subtotal
            sheet.Cells[10, 0].PutValue("Subtotal");

            // Freeze rows up to the subtotal row (rows 1‑11)
            // Parameters: totalRows, totalColumns, frozenRows, frozenColumns
            // Using 0 for totalRows/totalColumns means the whole sheet.
            sheet.FreezePanes(0, 0, 11, 0);

            // Save the workbook
            workbook.Save("SubtotalsAndFreeze.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
