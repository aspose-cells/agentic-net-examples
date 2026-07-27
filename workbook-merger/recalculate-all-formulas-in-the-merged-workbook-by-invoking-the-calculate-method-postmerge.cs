// Title: Merge Excel workbooks and recalculate formulas using Aspose.Cells for .NET (C#)
// Description: Loads two Excel files, copies every worksheet from the second workbook into the first with Workbook.Worksheets.AddCopy, runs Workbook.CalculateFormula to update all formulas, and saves the merged workbook. Includes basic file‑existence checks and exception handling.
// Keywords: Aspose.Cells merge workbooks C# | Workbook.CalculateFormula | AddCopy worksheet Aspose | recalculate formulas after merge | combine Excel files .NET | Excel workbook consolidation | Aspose.Cells error handling
// Common Searches: how to merge two Excel workbooks with Aspose.Cells | C# recalculate formulas after merging worksheets | Aspose.Cells AddCopy example | merge Excel files and refresh calculations .NET | Workbook.CalculateFormula usage
// Developer Intent: Combine multiple Excel files into one workbook and ensure every formula reflects the merged data.
// Use Cases: Consolidate monthly departmental reports into a single workbook before distribution. | Merge a master template with a data‑driven workbook and automatically refresh totals and percentages. | Automate the aggregation of quarterly financial statements and recompute summary calculations for a final audit file.
// AI Prompts: Write C# code that merges several Excel workbooks with Aspose.Cells and calls Workbook.CalculateFormula on the result. | Explain how Workbook.CalculateFormula works after adding worksheets via AddCopy. | Provide robust error‑handling patterns for missing source files and save failures when merging workbooks with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads two Excel files, copies every worksheet from the second workbook into the first with Workbook.Worksheets.AddCopy, runs Workbook.CalculateFormula to update all formulas, and saves the merged workbook. Includes basic file‑existence checks and exception handling.
class MergeAndCalculate
{
    static void Main()
    {
        try
        {
            const string source1Path = "source1.xlsx";
            const string source2Path = "source2.xlsx";
            const string outputPath = "merged_output.xlsx";

            // Verify that source files exist to avoid FileNotFoundException
            if (!File.Exists(source1Path))
                throw new FileNotFoundException($"The file '{source1Path}' was not found.");
            if (!File.Exists(source2Path))
                throw new FileNotFoundException($"The file '{source2Path}' was not found.");

            // Load the first workbook (creates the workbook instance)
            Workbook mergedWorkbook = new Workbook(source1Path);

            // Load the second workbook to be merged
            Workbook workbookToMerge = new Workbook(source2Path);

            // Merge: copy each worksheet from the second workbook into the first one
            foreach (Worksheet sheet in workbookToMerge.Worksheets)
            {
                // Add a copy of the worksheet to the merged workbook using the sheet name
                mergedWorkbook.Worksheets.AddCopy(sheet.Name);
            }

            // Recalculate all formulas in the merged workbook
            mergedWorkbook.CalculateFormula();

            // Save the merged and recalculated workbook
            mergedWorkbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
