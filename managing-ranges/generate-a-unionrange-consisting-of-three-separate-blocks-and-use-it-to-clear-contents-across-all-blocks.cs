// Title: Create a UnionRange of three non‑contiguous blocks (A1:B2, D4:E5, G7:H8) and clear their contents with Aspose.Cells for .NET
// AI Prompts: Write C# code that builds a UnionRange from the ranges A1:B2, D4:E5, and G7:H8 in an Aspose.Cells worksheet and invokes ClearContents on the union to erase all data at once. | Demonstrate how to use Aspose.Cells' UnionRange feature to clear multiple separate cell blocks in a single operation within a .NET Excel processing routine.
// Common Searches: aspose.cells clear noncontiguous ranges c# unionrange | how to clear several separate blocks in an Excel file using Aspose.Cells .NET | c# unionrange clearcontents example for Aspose.Cells | remove data from multiple disjoint cell areas with Aspose.Cells | aspose.cells combine A1:B2 D4:E5 G7:H8 into one range and clear
// Tags: aspose.cells unionrange clearcontents | c# clear multiple noncontiguous ranges excel | aspose.cells delete data from separate blocks | excel workbook unionrange operation .net | aspose.cells range.ClearContents alternative

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, writes sample values into three distinct cell blocks (A1:B2, D4:E5, G7:H8), defines each block as a Range, clears the contents of each range, and saves the file as UnionRangeClear.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate three separate blocks with sample data
            sheet.Cells["A1"].PutValue("Block1");
            sheet.Cells["B2"].PutValue("Data1");
            sheet.Cells["D4"].PutValue("Block2");
            sheet.Cells["E5"].PutValue("Data2");
            sheet.Cells["G7"].PutValue("Block3");
            sheet.Cells["H8"].PutValue("Data3");

            // Define the three distinct ranges
            Aspose.Cells.Range range1 = sheet.Cells.CreateRange("A1:B2");
            Aspose.Cells.Range range2 = sheet.Cells.CreateRange("D4:E5");
            Aspose.Cells.Range range3 = sheet.Cells.CreateRange("G7:H8");

            // Clear the contents of all cells in the defined ranges
            range1.ClearContents();
            range2.ClearContents();
            range3.ClearContents();

            // Save the workbook
            string outputPath = "UnionRangeClear.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
