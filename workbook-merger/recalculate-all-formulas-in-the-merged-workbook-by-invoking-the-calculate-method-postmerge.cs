// Title: Recalculate all formulas in a merged Excel workbook using Aspose.Cells CalculateFormula (C#)
// Common Searches: Aspose.Cells C# recalculate formulas after merging multiple .xlsx files | How to refresh all formulas in a merged workbook with Aspose.Cells CalculateFormula | Merge several Excel workbooks and evaluate formulas programmatically in .NET | C# code to combine workbooks and recalculate all formulas using Aspose.Cells | CalculateFormula method on a workbook that contains copied worksheets
// Tags: Aspose.Cells CalculateFormula merged workbook | C# merge Excel worksheets with formula refresh | recalculate formulas after workbook copy Aspose.Cells | merged workbook formula evaluation .NET | Excel file merge and formula recalculation using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeAndCalculate
{
    // The example loads multiple source Excel files, creates an empty Workbook, copies each worksheet from the source workbooks into the new workbook, calls mergedWorkbook.CalculateFormula() to recalculate every formula across all merged sheets, and saves the result as MergedAndCalculated.xlsx while handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            // List of source workbook file paths to be merged
            List<string> sourceFiles = new List<string>
            {
                "Source1.xlsx",
                "Source2.xlsx",
                // Add more source files as needed
            };

            // Create a new workbook that will hold the merged content
            Workbook mergedWorkbook = new Workbook();

            // Remove the default empty worksheet created by the constructor
            mergedWorkbook.Worksheets.Clear();

            // Iterate through each source workbook and copy its worksheets into the merged workbook
            foreach (string filePath in sourceFiles)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the source workbook
                    Workbook sourceWorkbook = new Workbook(filePath);

                    // Copy each worksheet from the source workbook
                    foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                    {
                        // Add a copy of the worksheet to the merged workbook using its name
                        mergedWorkbook.Worksheets.AddCopy(sheet.Name);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            try
            {
                // Recalculate all formulas in the merged workbook
                mergedWorkbook.CalculateFormula();

                // Save the merged and recalculated workbook
                mergedWorkbook.Save("MergedAndCalculated.xlsx");
                Console.WriteLine("Merged workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving merged workbook: {ex.Message}");
            }
        }
    }
}
