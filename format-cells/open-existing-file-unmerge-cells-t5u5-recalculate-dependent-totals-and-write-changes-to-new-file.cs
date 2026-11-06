// Title: Unmerge cells T5:U5 and recalculate dependent formulas in an existing Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unmerge merged cells T5:U5 in an existing workbook | Recalculate all formulas after modifying a range with Aspose.Cells .NET | Save changes to a new Excel file after unmerging cells using Aspose.Cells | Check if input.xlsx exists before processing with Aspose.Cells in C# | Programmatically create and unmerge a specific range with Aspose.Cells
// Tags: unmerge cells T5:U5 Aspose.Cells | recalculate dependent formulas Aspose.Cells | save workbook as new Excel file Aspose.Cells | load existing Excel workbook Aspose.Cells | create range T5:U5 Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, creates the range T5:U5 and calls UnMerge to split any merged cells, recalculates all formulas to update dependent totals, ensures the output directory is present, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Unmerge the range T5:U5 (zero‑based indices: T=19, U=20, row 5 = 4)
            AsposeRange mergedRange = sheet.Cells.CreateRange("T5:U5");
            mergedRange.UnMerge(); // UnMerge works even if the range is not merged

            // Recalculate all formulas to update dependent totals
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
