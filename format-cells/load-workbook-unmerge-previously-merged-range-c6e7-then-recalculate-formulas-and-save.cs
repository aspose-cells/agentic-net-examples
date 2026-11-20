// Title: C# example: Unmerge range C6:E7, recalculate all formulas, and save the workbook with Aspose.Cells
// Common Searches: Aspose.Cells C# unmerge cells C6:E7 and refresh formulas | How to programmatically unmerge a merged range in an Excel file using Aspose.Cells .NET | Recalculate workbook formulas after unmerging cells with Aspose.Cells for .NET | Save workbook after unmerging a specific range using Aspose.Cells C# | Example code to unmerge a range and recalculate formulas in Aspose.Cells
// Tags: Aspose.Cells unmerge range C6:E7 | Aspose.Cells recalculate formulas .NET | Aspose.Cells save workbook after unmerge | C# Excel unmerge cells example | Aspose.Cells workbook manipulation

using Aspose.Cells;
using System;
using System.IO;

// This C# program loads an existing Excel workbook (or creates one if missing), unmerges the merged range C6:E7, recalculates all formulas in the workbook, and saves the updated file as output.xlsx using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the input file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook();
                // Example merged range to demonstrate unmerge (C6:E7)
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells.Merge(5, 2, 2, 3);
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Unmerge the range C6:E7
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("C6:E7");
            mergedRange.UnMerge();

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
