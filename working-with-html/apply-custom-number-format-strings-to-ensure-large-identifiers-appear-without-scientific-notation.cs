// Title: Use Aspose.Cells in C# to apply a custom numeric format that prevents scientific notation for large integer IDs
// AI Prompts: Write C# code that inserts long integer identifiers into an Excel worksheet with Aspose.Cells and applies the custom format string "0" to each cell so the full value is displayed. | Demonstrate how to get a cell's style, set its Custom property to "0", and save the workbook to keep large numbers from appearing in exponential notation.
// Common Searches: aspnet c# aspose.cells prevent scientific notation for big numbers in Excel | apply custom format "0" to cells using Aspose.Cells C# | display 15‑digit IDs in Excel without exponential format with Aspose.Cells | set cell style custom format Aspose.Cells to preserve large integer values
// Tags: cell format "0" Aspose.Cells C# | prevent scientific notation Excel Aspose.Cells | set cell style Aspose.Cells C# | write large integer IDs Aspose.Cells | save workbook as xlsx Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a workbook, writes several large integer IDs to cells, applies the custom numeric format "0" to each cell to suppress scientific notation, and saves the file as LargeIdentifiers.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample large identifiers that could be shown in scientific notation
        long[] largeIds = new long[]
        {
            123456789012345,
            987654321098765432,
            1234567890123456789
        };

        // Write each identifier to a cell and apply a custom number format to force full display
        for (int i = 0; i < largeIds.Length; i++)
        {
            Cell cell = sheet.Cells[i, 0];
            cell.PutValue(largeIds[i]);               // store the numeric value

            // Apply custom format "0" to prevent scientific notation
            Style style = cell.GetStyle();
            style.Custom = "0";
            cell.SetStyle(style);
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("LargeIdentifiers.xlsx", SaveFormat.Xlsx);
    }
}
