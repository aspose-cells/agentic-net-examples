// Title: Offset a Range in Aspose.Cells for .NET – Shift D4:F10 by 3 Rows and 2 Columns (C#)
// Description: C# example that creates a workbook with Aspose.Cells, defines the range D4:F10, obtains an offset range three rows down and two columns to the right using Range.GetOffset, writes a value to the new top‑left cell, prints both range addresses, and saves the file as OffsetRangeDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Range.GetOffset | offset range example | shift rows columns | D4:F10 | Excel automation | Aspose.Cells workbook
// Common Searches: Aspose.Cells GetOffset C# example | how to offset a range in Aspose.Cells | shift range D4:F10 three rows two columns | Range.GetOffset usage .NET | offset range Excel library C#
// Developer Intent: Create a new range that is positioned a specific number of rows and columns away from an existing range.
// Use Cases: Generate a dynamic data entry block that moves relative to a template area. | Copy formulas, styles, or validation to a region offset from a source range. | Build multi‑section reports by repeatedly offsetting a base range for each section.
// AI Prompts: Write C# code with Aspose.Cells to offset a given range by N rows and M columns and insert a value in the offset range's first cell. | Explain how Range.GetOffset calculates the address of the new range and what limits apply. | Provide error‑handling logic for offset operations that might exceed worksheet boundaries.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// C# example that creates a workbook with Aspose.Cells, defines the range D4:F10, obtains an offset range three rows down and two columns to the right using Range.GetOffset, writes a value to the new top‑left cell, prints both range addresses, and saves the file as OffsetRangeDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create the original range D4:F10
            AsposeRange originalRange = cells.CreateRange("D4", "F10");

            // Get the offset range: 3 rows down and 2 columns right
            AsposeRange offsetRange = originalRange.GetOffset(3, 2);

            // Put a value in the top-left cell of the offset range
            offsetRange[0, 0].PutValue("OffsetStart");

            // Output the addresses for verification
            Console.WriteLine("Original Range Address: " + originalRange.Address);
            Console.WriteLine("Offset Range Address: " + offsetRange.Address);

            // Save the workbook
            workbook.Save("OffsetRangeDemo.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
