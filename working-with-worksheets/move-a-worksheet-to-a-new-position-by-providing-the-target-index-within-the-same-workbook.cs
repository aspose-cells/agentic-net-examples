// Title: Move a worksheet to a specific position within the same workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, validates source and destination indices, moves the worksheet with Worksheet.MoveTo, and saves the updated workbook. | Create a reusable C# method that takes a workbook path, source index, and target index, reorders the sheet using Aspose.Cells, and returns the path of the saved file. | Write a console application example that demonstrates moving the first worksheet to the third position in an Excel file with Aspose.Cells for .NET.
// Common Searches: aspnet move worksheet to third position Aspose.Cells example | c# Aspose.Cells reorder sheets by index | how to use Worksheet.MoveTo to change sheet order in .xlsx | validate worksheet index range before moving with Aspose.Cells C# | move sheet from index 0 to index 2 using Aspose.Cells .NET
// Tags: Worksheet.MoveTo method Aspose.Cells | reorder Excel worksheets .NET | check sheet index bounds C# | save reordered workbook Aspose | load .xlsx file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads "input.xlsx", ensures the source and target worksheet indices are valid, moves the worksheet at the source index to the target index using Worksheet.MoveTo, and saves the reordered workbook as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Index of the worksheet you want to move (0‑based)
            int sourceIndex = 0;   // e.g., first worksheet

            // Target index where the worksheet should be placed (0‑based)
            int targetIndex = 2;   // e.g., move to third position

            // Ensure source index is within range
            if (sourceIndex < 0 || sourceIndex >= workbook.Worksheets.Count)
            {
                Console.WriteLine("Error: Source index is out of range.");
                return;
            }

            // Ensure target index is within the allowed range (0 to Count)
            if (targetIndex < 0 || targetIndex > workbook.Worksheets.Count)
            {
                Console.WriteLine("Error: Target index is out of range.");
                return;
            }

            // Move the worksheet using the Worksheet.MoveTo method
            workbook.Worksheets[sourceIndex].MoveTo(targetIndex);

            // Save the workbook with the updated order
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
