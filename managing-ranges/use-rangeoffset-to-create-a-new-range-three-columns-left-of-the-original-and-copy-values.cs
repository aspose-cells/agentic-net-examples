// Title: Copy values to a range three columns left using Range.Offset in Aspose.Cells for .NET
// Description: Demonstrates creating a workbook, defining a source range, offsetting it three columns left with Range.GetOffset(0, -3), copying values via CopyValue, and saving the file.
// Keywords: Aspose.Cells Range.Offset example | CopyValue with offset range | C# Aspose.Cells move data left | Range.GetOffset negative column | Aspose.Cells .NET tutorial
// Common Searches: Aspose.Cells copy range three columns left | Range.GetOffset usage C# | How to shift a range left in Aspose.Cells | CopyValue offset range example | Aspose.Cells offset range tutorial
// Developer Intent: Create an offset range three columns left of an existing range and copy the original values into it.
// Use Cases: Generate a summary column adjacent to a data table by shifting values left. | Create a side‑by‑side copy of a table for comparison in a report. | Populate legacy columns with the same data as a newly added table.
// AI Prompts: Write C# code that uses Aspose.Cells to copy a range to a location three columns left using Range.GetOffset and CopyValue. | Explain how Range.GetOffset handles negative column offsets and how to match the size of the offset range with the source range. | Provide a step‑by‑step guide to create a source range, offset it, copy its values, and save the workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsOffsetExample
{
    // Demonstrates creating a workbook, defining a source range, offsetting it three columns left with Range.GetOffset(0, -3), copying values via CopyValue, and saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate a source range (e.g., D5:F7) with sample data
                AsposeRange sourceRange = cells.CreateRange("D5", "F7");
                for (int i = 0; i < sourceRange.RowCount; i++)
                {
                    for (int j = 0; j < sourceRange.ColumnCount; j++)
                    {
                        sourceRange[i, j].PutValue($"R{i + 5}C{j + 4}");
                    }
                }

                // Get a new range that is three columns to the left of the source range
                // Row offset = 0 (same rows), Column offset = -3 (three columns left)
                AsposeRange offsetRange = sourceRange.GetOffset(0, -3);

                // Copy the values from the source range to the offset range
                offsetRange.CopyValue(sourceRange);

                // Save the workbook
                string outputPath = "OffsetCopyDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
