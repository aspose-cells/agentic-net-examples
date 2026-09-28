// Title: Copy a formula‑filled range to a newly added worksheet while keeping the workbook’s calculation mode unchanged using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to copy a range that contains formulas from one sheet to a newly created sheet without altering the workbook’s calculation mode. | Add a new worksheet to an existing workbook and duplicate a table with formulas, ensuring the calculation setting (automatic or manual) stays the same after saving. | Adjust the copy operation so that the source worksheet’s calculation mode is preserved when transferring the range to another sheet with Aspose.Cells.
// Common Searches: Aspose.Cells copy range with formulas to another sheet keep calculation mode | C# add new worksheet and duplicate Excel table formulas using Aspose.Cells | preserve workbook calculation settings when copying formulas between worksheets Aspose.Cells | how to retain automatic calculation after copying cells in Aspose.Cells .NET
// Tags: copy formula range Aspose.Cells C# | add worksheet duplicate table Aspose.Cells | preserve calculation mode Aspose.Cells | workbook calculation settings .NET | duplicate Excel table with formulas Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, locates the "SourceSheet" worksheet, defines the A1:D10 range that contains formulas, adds a new worksheet named "CopiedTableSheet", creates a matching destination range, copies the source range (including formulas) to the new sheet, and saves the workbook while preserving the original calculation mode.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Get the source worksheet (replace with actual sheet name if different)
            var sourceSheet = workbook.Worksheets["SourceSheet"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Source worksheet 'SourceSheet' not found.");
                return;
            }

            // Define the range of the table to copy
            var sourceRange = sourceSheet.Cells.CreateRange("A1:D10");

            // Add a new worksheet to hold the copied table
            int newSheetIndex = workbook.Worksheets.Add();
            var newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "CopiedTableSheet";

            // Create a destination range in the new worksheet
            var destRange = newSheet.Cells.CreateRange("A1:D10");

            // Copy the source range (including formulas) to the destination range
            destRange.Copy(sourceRange);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
