// Title: Unmerge cells B2:C2 and transfer its original value to D2 using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells unmerge a merged range while preserving the original value | Copy value from merged cells B2:C2 to another cell after unmerging with Aspose.Cells | Read value from a merged Excel cell before unmerging using Aspose.Cells .NET | Programmatically unmerge cells in an Excel workbook and write original content to an adjacent cell in C# | Aspose.Cells keep data from merged cells when performing an unmerge operation
// Tags: Aspose.Cells unmerge range C# | retrieve merged cell value Aspose.Cells | assign cell value D2 Aspose.Cells | Excel workbook manipulation Aspose.Cells .NET | preserve original data after unmerge Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The C# program verifies the input file, loads it with Aspose.Cells, reads the value from the top‑left cell of the merged range B2:C2, unmerges that range, writes the captured value to cell D2, and saves the workbook as output.xlsx while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing spreadsheet
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the merged range B2:C2 (row 2, columns B and C)
            // Use fully qualified name to avoid ambiguity with System.Range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("B2:C2");

            // Retrieve the original value from the merged cell (top‑left cell B2)
            object originalValue = sheet.Cells["B2"].Value;

            // Unmerge the cells B2:C2
            mergedRange.UnMerge();

            // Copy the original value to the adjacent cell D2
            sheet.Cells["D2"].Value = originalValue;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
