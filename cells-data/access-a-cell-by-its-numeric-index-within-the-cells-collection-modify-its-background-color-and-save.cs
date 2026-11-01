// Title: Change a cell's background color by numeric row/column index using Aspose.Cells for .NET (C#) and save as XLSX
// AI Prompts: Use Aspose.Cells in C# to set a solid yellow background for a cell identified by zero‑based row and column indices and then save the workbook. | Create a new workbook, write a value to cell K1 using a numeric column index, apply a solid fill style, and persist the file with Aspose.Cells. | Retrieve a cell via sheet.Cells[row, column] in Aspose.Cells, modify its ForegroundColor and Pattern, and export the workbook to an .xlsx file.
// Common Searches: Aspose.Cells C# set background color of a cell using numeric row and column indices | How to apply solid fill to a specific cell by column index in Aspose.Cells for .NET | Save workbook after formatting a cell with Aspose.Cells C# example | Access cell K1 by zero‑based column index and change its fill color Aspose.Cells | C# Aspose.Cells change cell style and export to XLSX
// Tags: set cell background color Aspose.Cells C# | access cell by numeric index Aspose.Cells | apply solid fill style worksheet cell | save workbook as xlsx Aspose.Cells | modify cell style programmatically .NET

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example creates a new workbook, accesses cell K1 via zero‑based row and column indices, assigns a value, applies a solid yellow background style, and saves the workbook as ModifiedWorkbook.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access a cell by row and column (zero‑based). Example: first row, column 10 (K1)
            int columnIndex = 10; // zero‑based column index
            Cell cell = sheet.Cells[0, columnIndex];

            // Optional: put a value to visualize the cell
            cell.PutValue("Sample");

            // Modify the cell's background color
            Style style = cell.GetStyle();
            style.ForegroundColor = Color.Yellow;      // set desired background color
            style.Pattern = BackgroundType.Solid;      // ensure solid fill
            cell.SetStyle(style);

            // Define output file path
            string outputPath = "ModifiedWorkbook.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
