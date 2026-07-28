// Title: Create a Duplicate Range Reference with GetOffset(0,0) in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to obtain a second Range object that points to the same cells by calling GetOffset(0,0) on an existing range, modify a cell through the duplicate, and save the workbook.
// Keywords: Aspose.Cells GetOffset | duplicate range reference | C# Aspose.Cells example | range offset zero rows | Aspose.Cells range cloning
// Common Searches: Aspose.Cells GetOffset(0,0) usage | how to duplicate a range in Aspose.Cells | C# create reference to same cells Aspose.Cells | range offset zero rows example | clone range without copying data Aspose.Cells
// Developer Intent: Obtain a second Range object that references the same cells as an existing range by offsetting zero rows and zero columns.
// Use Cases: Pass the same cell block to multiple helper methods without recreating the range. | Demonstrate that changes via an offset range affect the original cells for validation purposes. | Maintain a clean API when a method requires a Range parameter but you want to reuse an existing range.
// AI Prompts: Write C# code that uses GetOffset(0,0) to create a duplicate range and then applies a style to the duplicate. | Show how to copy a range to a new location by offsetting non‑zero rows and columns with Aspose.Cells. | Explain the difference between GetOffset(0,0) and creating a new Range object when you need only a reference, not a copy.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to obtain a second Range object that points to the same cells by calling GetOffset(0,0) on an existing range, modify a cell through the duplicate, and save the workbook.
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

            // Create an initial range (A1:B2) and fill it with sample data
            AsposeRange original = cells.CreateRange("A1", "B2");
            for (int i = 0; i < original.RowCount; i++)
            {
                for (int j = 0; j < original.ColumnCount; j++)
                {
                    original[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Offset the range by zero rows and zero columns to obtain a duplicate reference
            AsposeRange duplicate = original.GetOffset(0, 0);

            // Modify a cell via the duplicate reference to prove it points to the same cells
            duplicate[0, 0].PutValue("Changed");

            // Save the workbook
            string outputPath = "OffsetZeroDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
