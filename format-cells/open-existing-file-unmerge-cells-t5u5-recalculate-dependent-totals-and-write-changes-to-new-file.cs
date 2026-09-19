// Title: Unmerge cells T5:U5 in an Excel workbook, recalculate dependent formulas, and save to a new file with Aspose.Cells for .NET (C#)
// AI Prompts: Load a workbook, unmerge the range T5:U5, run full formula calculation, and write the updated workbook to a different file using Aspose.Cells in C#. | Create C# code that opens input.xlsx, removes the merge on cells T5 through U5, triggers recalculation of all dependent formulas, and saves the result as output.xlsx with Aspose.Cells.
// Common Searches: asp.net unmerge cells T5 U5 using Aspose.Cells | recalculate all formulas after unmerging a range in Excel with Aspose.Cells C# | save modified Excel workbook to a new file after unmerge operation Aspose.Cells | how to refresh dependent totals when unmerging cells in Aspose.Cells .NET | C# example for unmerging a specific cell range and updating formulas with Aspose.Cells
// Tags: unmerge cell range Aspose.Cells | formula recalculation Aspose.Cells | save workbook as new file Aspose.Cells | Aspose.Cells range unmerge example | Excel cell merge handling .NET

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program loads "input.xlsx", accesses the first worksheet, unmerges the cells T5:U5, recalculates all formulas that depend on those cells, and saves the modified workbook as "output.xlsx", handling missing files and runtime exceptions.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Unmerge the range T5:U5 using a fully qualified Range type
            AsposeRange range = sheet.Cells.CreateRange("T5:U5");
            range.UnMerge();

            // Recalculate all formulas that depend on the unmerged cells
            workbook.CalculateFormula();

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
