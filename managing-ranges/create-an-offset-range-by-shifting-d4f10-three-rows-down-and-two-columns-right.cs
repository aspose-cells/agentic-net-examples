// Title: Shift range D4:F10 three rows down and two columns right using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to move a range down three rows and right two columns | Create a new range from D4:F10 with offset using Aspose.Cells .NET | Offset an existing Excel range programmatically with Aspose.Cells | C# code to shift a cell range by rows and columns in Aspose.Cells | Aspose.Cells example for creating a range based on another range's position
// Tags: Aspose.Cells offset range creation | CreateRange with row and column offset .NET | Excel range shift rows columns Aspose.Cells | Aspose.Cells range manipulation example | C# Aspose.Cells offset range D4:F10 | Aspose.Cells workbook range offset operation

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The sample creates a workbook, defines the original range D4:F10, computes an offset range three rows lower and two columns to the right, writes "Offset Start" into the first cell of the offset range, and saves the file as OffsetRangeDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the original range D4:F10 (zero‑based indices)
            // Row 3 (D4), Column 3 (D), 7 rows, 3 columns (D to F)
            AsposeRange originalRange = sheet.Cells.CreateRange(3, 3, 7, 3);

            // Calculate the offset range (3 rows down, 2 columns right)
            int offsetRows = 3;
            int offsetCols = 2;
            AsposeRange offsetRange = sheet.Cells.CreateRange(
                originalRange.FirstRow + offsetRows,
                originalRange.FirstColumn + offsetCols,
                originalRange.RowCount,
                originalRange.ColumnCount);

            // Put a value in the first cell of the offset range
            offsetRange[0, 0].PutValue("Offset Start");

            // Save the workbook
            string outputPath = "OffsetRangeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
