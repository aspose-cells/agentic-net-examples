// Title: Delete rows 30‑42, insert five blank rows, and save the worksheet as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that eliminates rows 30‑42, inserts five blank rows at the same index, and saves the workbook as a PDF. | Generate an example showing how to erase a specific row block, add new rows with default height, and convert the updated Excel file to PDF using Aspose.Cells in .NET.
// Common Searches: Aspose.Cells C# how to delete a range of rows and then add rows before exporting to PDF | remove rows 30-42 and add rows in Excel with Aspose.Cells, then save as PDF | C# Aspose.Cells example for row manipulation and PDF conversion
// Tags: Aspose.Cells row removal C# | Aspose.Cells row insertion default height C# | Aspose.Cells PDF export C# | modify worksheet rows Aspose.Cells C#

using Aspose.Cells;
using System;

// // Loads 'input.xlsx', eliminates rows 30‑42, inserts five new rows with default height at that position, and saves the result as 'output.pdf' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Delete rows 30 to 42 (1‑based). Zero‑based start index is 29, total rows = 13.
        sheet.Cells.DeleteRows(29, 13);

        // Insert five new rows at the same position with default height.
        sheet.Cells.InsertRows(29, 5);

        // Export the workbook to PDF.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
