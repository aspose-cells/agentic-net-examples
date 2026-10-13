// Title: Check that Cells.Subtotal raises a clear exception when the start row index is greater than the end row index in Aspose.Cells for .NET
// AI Prompts: Write a C# example that invokes Workbook.Worksheets[0].Cells.Subtotal with a CellArea where StartRow exceeds EndRow and captures the exception message. | Add an NUnit assertion to confirm that Cells.Subtotal throws an ArgumentException when the row range is reversed. | Log the full exception details from Cells.Subtotal when the CellArea has StartRow larger than EndRow and output the message to the console.
// Common Searches: Aspose.Cells Subtotal method error when start row is after end row | how to catch exception from Cells.Subtotal invalid row range in C# | validate CellArea parameters before calling Subtotal in Aspose.Cells | exception message for reversed startRow and endRow in Aspose.Cells Subtotal
// Tags: Aspose.Cells Subtotal invalid row range | C# Cells.Subtotal argument validation | CellArea startRow greater than endRow handling | Aspose.Cells exception handling for Subtotal | subtotal method error message Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates calling Cells.Subtotal with a CellArea where StartRow (5) is greater than EndRow (2), catches the resulting exception, and shows how to verify that a descriptive error is thrown.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("A");
            cells["B4"].PutValue(30);
            cells["A5"].PutValue("B");
            cells["B5"].PutValue(40);

            // Intentionally set startRow greater than endRow (zero‑based indices)
            int startRow = 5; // corresponds to Excel row 6
            int endRow = 2;   // corresponds to Excel row 3

            // Define the cell area for the subtotal operation.
            CellArea area = new CellArea
            {
                StartRow = startRow,
                StartColumn = 0,
                EndRow = endRow,
                EndColumn = 1
            };

            // Attempt to create a subtotal with invalid row range.
            // Subtotal(area, columnIndex, function, totalColumns, replace, pageBreak, summaryBelowData)
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, true, true);
            Console.WriteLine("No exception was thrown – validation failed.");
        }
        catch (Exception ex)
        {
            // Expected path: an informative exception should be thrown
            Console.WriteLine("Caught expected exception:");
            Console.WriteLine(ex.Message);
        }
    }
}
