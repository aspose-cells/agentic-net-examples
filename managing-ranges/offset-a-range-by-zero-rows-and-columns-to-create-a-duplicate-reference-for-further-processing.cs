// Title: Create a duplicate Range using GetOffset(0,0) in Aspose.Cells for .NET (C#)
// Description: This C# example builds a workbook, defines the range A1:B2, fills it with sample data, and then calls GetOffset(0,0) to obtain another Range object that points to the identical cells. Changes made through the new reference are reflected in the original range, and the workbook is saved as OffsetZeroDemo.xlsx.
// Keywords: Aspose.Cells, C#, GetOffset, identical range, same address, Excel automation, .NET, range reference, workbook example
// Common Searches: Aspose.Cells GetOffset zero offset identical range | How to obtain a second Range object that points to the same cells | C# Aspose.Cells range reference without cloning | GetOffset(0,0) effect on original range
// Developer Intent: Retrieve a Range instance that references the exact same cells as an existing range without creating a copy.
// Use Cases: Pass the zero‑offset Range to APIs that require a Range parameter while preserving the original variable. | Apply formatting, formulas, or data updates through the new reference and have them instantly appear in the source range. | Reuse the same Range object inside loops to avoid repeated range‑creation overhead.
// AI Prompts: Write C# code that gets a zero‑offset Range with Aspose.Cells and sets a background color on it. | Show how to use the identical Range returned by GetOffset(0,0) to transfer values to another worksheet. | Compare GetOffset(0,0) with Range.CopyTo for scenarios where the same cell area must be accessed.

using System;
using Aspose.Cells;

// This C# example builds a workbook, defines the range A1:B2, fills it with sample data, and then calls GetOffset(0,0) to obtain another Range object that points to the identical cells. Changes made through the new reference are reflected in the original range, and the workbook is saved as OffsetZeroDemo.xlsx.
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

            // Create an initial range (A1:B2)
            Aspose.Cells.Range originalRange = cells.CreateRange("A1", "B2");

            // Fill the original range with sample data
            for (int i = 0; i < originalRange.RowCount; i++)
            {
                for (int j = 0; j < originalRange.ColumnCount; j++)
                {
                    originalRange[i, j].PutValue($"R{i}C{j}");
                }
            }

            // Obtain a duplicate reference by offsetting zero rows and zero columns
            Aspose.Cells.Range duplicateRange = originalRange.GetOffset(0, 0);

            // Display addresses to confirm they are identical
            Console.WriteLine("Original range address: " + originalRange.Address);
            Console.WriteLine("Duplicate range address: " + duplicateRange.Address);

            // Demonstrate that changes via the duplicate affect the original range
            duplicateRange[0, 0].PutValue("Modified");

            // Save the workbook
            workbook.Save("OffsetZeroDemo.xlsx");
            Console.WriteLine("Workbook saved as OffsetZeroDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
