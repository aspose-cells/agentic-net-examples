// Title: Merge two Excel workbooks and recalculate all formulas with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens two .xlsx files, copies every worksheet into a new workbook, removes the initial empty sheet, calls Workbook.CalculateFormula, and saves the merged file. | Write a .NET program using Aspose.Cells to combine worksheets from multiple workbooks and automatically recalculate every formula before exporting. | Create an example that demonstrates merging Excel workbooks, invoking CalculateFormula on the combined workbook, and writing the result as merged.xlsx.
// Common Searches: Aspose.Cells C# merge two workbooks and recalculate formulas | How to use Workbook.CalculateFormula after copying worksheets with Aspose.Cells | C# code to combine Excel files and update all formulas using Aspose.Cells | Remove default sheet and merge worksheets in Aspose.Cells before saving | Recalculate formulas in a merged workbook using Aspose.Cells .NET
// Tags: Aspose.Cells merge worksheets C# | Workbook.CalculateFormula after merge | copy worksheets between workbooks Aspose.Cells | clear default worksheet Aspose.Cells | merge Excel files recalculate formulas .NET

using System;
using System.IO;
using Aspose.Cells;

// // Loads two Excel files, copies all their worksheets into a new workbook after clearing the default sheet, recalculates every formula with CalculateFormula, and saves the merged workbook as merged.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Verify source files exist before loading
            const string sourcePath1 = "source1.xlsx";
            const string sourcePath2 = "source2.xlsx";

            if (!File.Exists(sourcePath1))
                throw new FileNotFoundException($"Source file not found: {sourcePath1}");
            if (!File.Exists(sourcePath2))
                throw new FileNotFoundException($"Source file not found: {sourcePath2}");

            // Load the workbooks that need to be merged
            Workbook wb1 = new Workbook(sourcePath1);
            Workbook wb2 = new Workbook(sourcePath2);

            // Create a new workbook that will contain the merged worksheets
            Workbook merged = new Workbook();

            // Remove the default empty worksheet created by the constructor
            merged.Worksheets.Clear();

            // Copy all worksheets from the first source workbook
            foreach (Worksheet ws in wb1.Worksheets)
            {
                merged.Worksheets.AddCopy(ws.Name);
            }

            // Copy all worksheets from the second source workbook
            foreach (Worksheet ws in wb2.Worksheets)
            {
                merged.Worksheets.AddCopy(ws.Name);
            }

            // Recalculate all formulas in the merged workbook
            merged.CalculateFormula();

            // Save the merged workbook to a file
            const string outputPath = "merged.xlsx";
            merged.Save(outputPath);
            Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
