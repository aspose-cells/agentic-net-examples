// Title: Validate #REF! Errors After Merging Excel Workbooks with Aspose.Cells for .NET
// Description: C# sample that loads a primary workbook, copies worksheets from additional .xlsx files, recalculates all formulas, scans every cell for #REF! errors, reports any issues, and saves the consolidated workbook.
// Keywords: Aspose.Cells | C# workbook merge | Excel #REF! error detection | formula recalculation | Excel file consolidation | .NET Excel processing | reference error validation
// Common Searches: Aspose.Cells merge multiple workbooks and check for #REF! errors | C# code to combine Excel files and validate formulas | detect reference errors after copying worksheets with Aspose | how to recalculate formulas after Excel workbook merge .NET | validate merged Excel workbook for #REF! using Aspose.Cells
// Developer Intent: Combine several Excel files into one workbook and automatically verify that no #REF! reference errors remain after formulas are recalculated.
// Use Cases: Consolidate monthly financial statements from separate departments while ensuring all calculations stay intact. | Aggregate regional sales reports into a master workbook and abort the process if any reference errors appear. | Integrate workbook merging into a CI/CD pipeline, failing the build when #REF! errors are detected post‑merge.
// AI Prompts: Create C# code with Aspose.Cells that merges an array of workbook paths, recalculates formulas, and lists any #REF! errors found. | Modify the sample to write #REF! error details to a log file and return a boolean indicating validation success. | Explain how to extend the error‑checking loop to capture other Excel errors such as #DIV/0! and #VALUE! using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRefErrorCheck
{
    // C# sample that loads a primary workbook, copies worksheets from additional .xlsx files, recalculates all formulas, scans every cell for #REF! errors, reports any issues, and saves the consolidated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths of workbooks to be merged
                string[] workbookFiles = { "Workbook1.xlsx", "Workbook2.xlsx", "Workbook3.xlsx" };

                // Verify that the first workbook exists
                if (!File.Exists(workbookFiles[0]))
                {
                    Console.WriteLine($"Error: File not found - {workbookFiles[0]}");
                    return;
                }

                // Load the first workbook as the base workbook
                Workbook mergedWorkbook;
                try
                {
                    mergedWorkbook = new Workbook(workbookFiles[0]);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load '{workbookFiles[0]}': {ex.Message}");
                    return;
                }

                // Merge remaining workbooks by copying their worksheets into the base workbook
                for (int i = 1; i < workbookFiles.Length; i++)
                {
                    string filePath = workbookFiles[i];

                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Warning: File not found - {filePath}. Skipping.");
                        continue;
                    }

                    Workbook wbToMerge;
                    try
                    {
                        wbToMerge = new Workbook(filePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to load '{filePath}': {ex.Message}. Skipping.");
                        continue;
                    }

                    // Copy each worksheet from the source workbook into the merged workbook
                    foreach (Worksheet srcSheet in wbToMerge.Worksheets)
                    {
                        // AddCopy expects the source sheet name
                        mergedWorkbook.Worksheets.AddCopy(srcSheet.Name);
                    }
                }

                // Ensure all formulas are evaluated after merging
                mergedWorkbook.CalculateFormula();

                // Validate #REF! errors
                bool hasRefError = false;

                foreach (Worksheet sheet in mergedWorkbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Enumerate all cells in the worksheet
                    foreach (Cell cell in cells)
                    {
                        // Check if the cell contains an error and specifically a #REF! error
                        if (cell.Type == CellValueType.IsError && cell.StringValue == "#REF!")
                        {
                            hasRefError = true;
                            Console.WriteLine($"#REF! error found in sheet '{sheet.Name}', cell {cell.Name}");
                        }
                    }
                }

                if (!hasRefError)
                {
                    Console.WriteLine("Validation passed: No #REF! errors exist in any cell.");
                }

                // Save the merged workbook
                try
                {
                    mergedWorkbook.Save("MergedWorkbook_Output.xlsx");
                    Console.WriteLine("Merged workbook saved as 'MergedWorkbook_Output.xlsx'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save merged workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
