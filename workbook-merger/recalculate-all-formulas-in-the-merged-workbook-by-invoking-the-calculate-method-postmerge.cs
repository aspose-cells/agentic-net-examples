// Title: Merge Excel workbooks and recalculate formulas with Aspose.Cells for .NET
// Description: Loads a target and a source workbook, copies each worksheet from the source into the target using AddCopy, runs Workbook.CalculateFormula to update every formula, and saves the merged file. Includes error handling for missing files.
// Keywords: Aspose.Cells | merge workbooks | calculate formulas | C# .NET | AddCopy | Workbook.CalculateFormula | copy worksheets | Excel file consolidation | formula recalculation | merged workbook
// Common Searches: Aspose.Cells merge workbooks C# | CalculateFormula after adding worksheets | Copy worksheets between workbooks Aspose.Cells | Recalculate all formulas in merged Excel file .NET | Combine two Excel files with Aspose.Cells
// Developer Intent: Combine two Excel workbooks into one and refresh every formula in the resulting file.
// Use Cases: Consolidate monthly reports into a master workbook while ensuring totals and derived values are up‑to‑date. | Merge scenario sheets from separate financial models and automatically recalculate dependent calculations. | Automate the creation of a final report by merging a template workbook with data workbooks, delivering accurate formula results.
// AI Prompts: Write C# code that merges multiple Excel workbooks using Aspose.Cells and invokes CalculateFormula on the combined workbook. | Suggest robust error‑handling patterns for loading, merging, and recalculating formulas with Aspose.Cells in a .NET application. | Explain how to recalculate formulas only on selected worksheets after a workbook merge using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads a target and a source workbook, copies each worksheet from the source into the target using AddCopy, runs Workbook.CalculateFormula to update every formula, and saves the merged file. Includes error handling for missing files.
class Program
{
    static void Main()
    {
        try
        {
            const string targetPath = "source1.xlsx";
            const string sourcePath = "source2.xlsx";
            const string outputPath = "merged_output.xlsx";

            // Verify that input files exist to avoid FileNotFoundException
            if (!File.Exists(targetPath))
                throw new FileNotFoundException($"Target workbook not found: {targetPath}");
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

            // Load the primary workbook (the one that will receive the merged content)
            Workbook targetWorkbook = new Workbook(targetPath);

            // Load the secondary workbook whose worksheets will be merged into the target
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Iterate through each worksheet in the source workbook and add a copy to the target workbook
            foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
            {
                // AddCopy expects the name of the worksheet to copy
                targetWorkbook.Worksheets.AddCopy(sourceSheet.Name);
            }

            // Recalculate all formulas in the combined workbook
            targetWorkbook.CalculateFormula();

            // Save the merged workbook with updated formula results
            targetWorkbook.Save(outputPath);
            Console.WriteLine($"Workbooks merged successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
