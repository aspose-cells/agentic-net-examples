// Title: How to create a UnionRange of multiple non‑contiguous blocks and clear their contents with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# clear contents of noncontiguous range A1:B2 D4:E5 G7:H8 | Create a union range with several areas using Aspose.Cells CreateRange method | C# Aspose.Cells example to clear cells in a combined range | How to clear multiple separate cell blocks in one operation with Aspose.Cells | Aspose.Cells .NET code for clearing contents of several cell blocks at once
// Tags: unionrange api Aspose.Cells | multi-section range creation C# | noncontiguous cell block handling | aggregate range manipulation .NET | excel workbook cell clearing

using System;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first Worksheet, defines a UnionRange that combines the blocks A1:B2, D4:E5, and G7:H8 using the CreateRange method with a comma‑separated address list, clears the contents of all cells in that combined range, and saves the workbook as UnionRangeClear.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a range that combines multiple blocks: A1:B2, D4:E5, G7:H8
            // Use the Aspose.Cells.Range type explicitly to avoid conflict with System.Range
            Aspose.Cells.Range unionRange = sheet.Cells.CreateRange("A1:B2,D4:E5,G7:H8");

            // Clear the contents of all cells in the combined range
            unionRange.ClearContents();

            // Save the workbook
            string outputPath = "UnionRangeClear.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
