// Title: Unmerge cells C6:E7, recalculate all formulas, and save the workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to open an Excel file, unmerge the range C6:E7, recalculate every formula, and save the result to a new file. | Show how to combine Aspose.Cells Range.UnMerge and Workbook.CalculateFormula methods in a .NET application. | Provide a C# example that loads a workbook, removes a merged region, forces formula evaluation, and writes the updated workbook back to disk.
// Common Searches: Aspose.Cells C# unmerge merged cells C6:E7 and recalculate formulas before saving | How to programmatically unmerge a specific range and refresh formulas with Aspose.Cells for .NET | Recalculate workbook formulas after unmerging cells using Aspose.Cells in C# | Save Excel file after unmerging cells with Aspose.Cells .NET API
// Tags: Aspose.Cells unmerge range C6:E7 | Aspose.Cells recalculate formulas | Aspose.Cells save modified workbook | Aspose.Cells Range.UnMerge example | Aspose.Cells Workbook.CalculateFormula usage

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads an existing Excel workbook, accesses the first worksheet, creates a Range object for cells C6:E7, calls UnMerge to split the merged cells, triggers a full formula recalculation with CalculateFormula, and saves the updated workbook to a new file, handling any errors that may occur.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Unmerge the previously merged range C6:E7 using a Range object
            AsposeRange mergedRange = sheet.Cells.CreateRange("C6:E7");
            mergedRange.UnMerge();

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
