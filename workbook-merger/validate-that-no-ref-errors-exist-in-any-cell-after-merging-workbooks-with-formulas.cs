// Title: C# code to merge multiple Excel workbooks with Aspose.Cells and validate that no #REF! errors remain
// Common Searches: C# Aspose.Cells how to identify #REF! cells after merging several Excel files | verify formula references after copying worksheets with Aspose.Cells .NET | find broken references in a merged workbook using Aspose.Cells API | recalculate all formulas and report errors after workbook merge in C# | scan used range for error values after combining Excel workbooks with Aspose.Cells
// Tags: merge worksheets Aspose.Cells .NET | detect #REF! errors after workbook merge | recalculate formulas merged workbook C# | validate cell error values Aspose.Cells | scan used range for reference errors

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookMergeValidation
{
    // The example loads a set of Excel files, creates an empty workbook, copies each source worksheet into it, recalculates all formulas, then iterates through the used range of every worksheet to locate cells whose type is an error and whose string value equals "#REF!". It reports any found reference errors, confirms validation when none are present, and saves the merged workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // List of source workbook file paths to be merged
                string[] sourceFiles = { "Workbook1.xlsx", "Workbook2.xlsx", "Workbook3.xlsx" };

                // Create a new empty workbook that will hold the merged content
                Workbook mergedWorkbook = new Workbook();

                // Remove the default empty worksheet created by Aspose.Cells
                if (mergedWorkbook.Worksheets.Count > 0)
                {
                    mergedWorkbook.Worksheets.RemoveAt(0);
                }

                // Load each source workbook and copy its worksheets into the merged workbook
                foreach (string filePath in sourceFiles)
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: {filePath}");
                        continue; // Skip missing files
                    }

                    try
                    {
                        // Load source workbook
                        Workbook sourceWorkbook = new Workbook(filePath);

                        // Copy each worksheet from the source workbook
                        foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                        {
                            // Add a copy of the worksheet to the merged workbook using the sheet name
                            mergedWorkbook.Worksheets.AddCopy(sheet.Name);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error loading or copying from '{filePath}': {ex.Message}");
                    }
                }

                // Recalculate all formulas in the merged workbook to ensure up‑to‑date results
                mergedWorkbook.CalculateFormula();

                // Validate that no cell contains a #REF! error after merging
                List<string> refErrorCells = new List<string>();

                foreach (Worksheet sheet in mergedWorkbook.Worksheets)
                {
                    // Get the used range of the worksheet to limit iteration to populated cells
                    var usedRange = sheet.Cells.MaxDisplayRange; // Aspose.Cells.Range

                    // Iterate through each cell in the used range
                    foreach (Cell cell in usedRange)
                    {
                        // Check if the cell contains an error value and specifically #REF!
                        if (cell.Type == CellValueType.IsError && cell.StringValue == "#REF!")
                        {
                            // Record the cell address with the #REF! error
                            refErrorCells.Add($"{sheet.Name}!{cell.Name}");
                        }
                    }
                }

                // Output validation result
                if (refErrorCells.Count == 0)
                {
                    Console.WriteLine("Validation passed: No #REF! errors found in the merged workbook.");
                }
                else
                {
                    Console.WriteLine("Validation failed: #REF! errors detected in the following cells:");
                    foreach (string address in refErrorCells)
                    {
                        Console.WriteLine(address);
                    }
                }

                // Save the merged workbook (optional)
                string outputPath = "MergedWorkbook.xlsx";
                try
                {
                    mergedWorkbook.Save(outputPath);
                    Console.WriteLine($"Merged workbook saved to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving merged workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
