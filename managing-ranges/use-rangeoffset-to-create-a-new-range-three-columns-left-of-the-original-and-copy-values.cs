// Title: Aspose.Cells .NET: Copy a range three columns left with Range.GetOffset
// Description: Demonstrates how to create a workbook, define a range (D1:F3), obtain an offset range three columns to the left (A1:C3) using GetOffset(0, -3), copy the original values with CopyValue, and save the file.
// Keywords: Aspose.Cells Range GetOffset | CopyValue offset range | C# Aspose.Cells example | move range left Aspose.Cells | CreateRange CopyValue .NET
// Common Searches: Aspose.Cells offset range left | Copy values to another range C# | GetOffset example Aspose.Cells | How to shift a range in Aspose.Cells | Range.CopyValue usage
// Developer Intent: Generate a new range positioned three columns left of an existing range and duplicate its values.
// Use Cases: Create a backup copy of a data block next to the original for quick comparison. | Populate a summary area by re‑using source data at a calculated offset. | Adjust a template layout by moving a table left without altering the source cells.
// AI Prompts: Show how to use Aspose.Cells Range.GetOffset to create a range three columns left and copy its values in C#. | Provide a C# code snippet that defines range D1:F3, offsets it by -3 columns, copies the data, and saves the workbook. | Explain the behavior of Range.CopyValue when copying between non‑overlapping ranges in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsOffsetCopyDemo
{
    // Demonstrates how to create a workbook, define a range (D1:F3), obtain an offset range three columns to the left (A1:C3) using GetOffset(0, -3), copy the original values with CopyValue, and save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Fill sample data in the original range (D1:F3)
                Cells cells = sheet.Cells;
                cells["D1"].PutValue("A");
                cells["E1"].PutValue("B");
                cells["F1"].PutValue("C");
                cells["D2"].PutValue(1);
                cells["E2"].PutValue(2);
                cells["F2"].PutValue(3);
                cells["D3"].PutValue(4);
                cells["E3"].PutValue(5);
                cells["F3"].PutValue(6);

                // Create the original range object
                AsposeRange originalRange = cells.CreateRange("D1:F3");

                // Get a new range that is three columns to the left of the original range
                // Row offset = 0 (same rows), Column offset = -3 (three columns left)
                AsposeRange offsetRange = originalRange.GetOffset(0, -3); // This will be A1:C3

                // Copy the values from the original range to the offset range
                offsetRange.CopyValue(originalRange);

                // Save the workbook to a file
                workbook.Save("OffsetCopyDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
