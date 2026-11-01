// Title: Recalculate formulas after merging worksheets with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells recalculate formulas after merging two workbooks | How to force formula recalculation in a merged Excel file using Aspose.Cells | Aspose.Cells CalculateFormula method after adding worksheets programmatically | Merge Excel workbooks and update all formulas with Aspose.Cells .NET | Recalculate all cells after workbook merge using Aspose.Cells API
// Tags: Aspose.Cells CalculateFormula after merging workbooks | addcopy worksheet recalc formulas C# | merged Excel workbook formula refresh Aspose.Cells | update all formulas post workbook merge .NET | Excel workbook merge with formula recalculation using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a primary and a secondary Excel workbook, copies each worksheet from the secondary into the primary using AddCopy, invokes CalculateFormula to recompute every formula across the merged workbook, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string primaryPath = "PrimaryWorkbook.xlsx";
            const string mergePath = "WorkbookToMerge.xlsx";
            const string outputPath = "MergedWorkbook.xlsx";

            // Verify input files exist
            if (!File.Exists(primaryPath))
            {
                Console.WriteLine($"Error: File not found - {primaryPath}");
                return;
            }
            if (!File.Exists(mergePath))
            {
                Console.WriteLine($"Error: File not found - {mergePath}");
                return;
            }

            // Load the primary workbook
            Workbook primaryWb = new Workbook(primaryPath);

            // Load the workbook to be merged
            Workbook mergeWb = new Workbook(mergePath);

            // Copy each worksheet from the merge workbook into the primary workbook
            foreach (Worksheet sheet in mergeWb.Worksheets)
            {
                // Add a copy of the source worksheet to the destination workbook using its name
                primaryWb.Worksheets.AddCopy(sheet.Name);
            }

            // Recalculate all formulas in the merged workbook
            primaryWb.CalculateFormula();

            // Save the merged workbook with recalculated formulas
            primaryWb.Save(outputPath);
            Console.WriteLine($"Merged workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
