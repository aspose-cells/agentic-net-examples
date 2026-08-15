// Title: Shift a Range in Aspose.Cells for .NET – Offset D4:F10 by 3 Rows and 2 Columns (C#)
// Description: Shows how to create a workbook, define the range D4:F10, apply GetOffset(3,2) to obtain the shifted range F7:H13, write sample values, and save the result as OffsetRangeDemo.xlsx.
// Keywords: Aspose.Cells | C# | GetOffset | offset range | shift range rows columns | D4:F10 | range address | Excel automation | CreateRange | Workbook save
// Common Searches: Aspose.Cells GetOffset example | How to offset a range in C# | Shift Excel range by rows and columns Aspose | Address of offset range D4:F10 | Aspose.Cells range manipulation
// Developer Intent: Create a new range that is the original D4:F10 moved three rows down and two columns to the right.
// Use Cases: Copy a data block to a new location while preserving its layout. | Place a summary table relative to source data by using an offset range. | Apply formulas or conditional formatting to a region that mirrors another range after shifting.
// AI Prompts: Write C# code using Aspose.Cells to offset range D4:F10 by 3 rows and 2 columns and display both addresses. | Provide an Aspose.Cells .NET example that creates an offset range, inserts sample values, and saves the workbook. | Explain how GetOffset calculates the new address and what happens when the offset exceeds worksheet boundaries.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Shows how to create a workbook, define the range D4:F10, apply GetOffset(3,2) to obtain the shifted range F7:H13, write sample values, and save the result as OffsetRangeDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create the original range D4:F10
            AsposeRange originalRange = cells.CreateRange("D4", "F10");

            // Shift the range 3 rows down and 2 columns right
            AsposeRange offsetRange = originalRange.GetOffset(3, 2);

            // Output the addresses of both ranges
            Console.WriteLine("Original Range Address: " + originalRange.Address);
            Console.WriteLine("Offset Range Address:   " + offsetRange.Address);

            // (Optional) Put sample values to verify the offset range
            originalRange[0, 0].PutValue("Original");
            offsetRange[0, 0].PutValue("Offset");

            // Save the workbook
            workbook.Save("OffsetRangeDemo.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
