// Title: Detect and Prevent #REF! Errors After Merging Workbooks with Formulas Using Aspose.Cells for .NET
// Description: C# sample that merges two Excel workbooks with Aspose.Cells, recalculates all formulas, scans every cell for #REF! error values, logs any broken references, and saves the consolidated file. Ideal for ensuring formula integrity after workbook consolidation.
// Keywords: Aspose.Cells merge workbooks | detect #REF errors C# | validate formulas after merge | Excel #REF detection .NET | workbook consolidation Aspose | C# Excel error scanning
// Common Searches: Aspose.Cells find #REF after merging workbooks | C# check for broken references in merged Excel file | how to validate formulas after workbook consolidation .NET | detect #REF! errors programmatically with Aspose.Cells | merge Excel files and ensure no reference errors
// Developer Intent: Confirm that no #REF! errors remain in any cell after merging workbooks and recalculating formulas.
// Use Cases: Combine departmental financial reports and automatically verify that all formulas resolve without broken references. | Consolidate template sheets into a master workbook while flagging any reference errors introduced by the merge. | Automate quarterly spreadsheet aggregation, guaranteeing formula integrity before distribution to stakeholders.
// AI Prompts: Create a reusable C# method that scans an Aspose.Cells Workbook for #REF! errors after CalculateFormula and returns a list of worksheet‑cell addresses. | Suggest an alternative merging approach with Aspose.Cells that preserves external links and minimizes #REF! occurrences. | Generate code to export detected #REF! error details to a CSV log instead of writing to the console.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeValidation
{
    // C# sample that merges two Excel workbooks with Aspose.Cells, recalculates all formulas, scans every cell for #REF! error values, logs any broken references, and saves the consolidated file. Ideal for ensuring formula integrity after workbook consolidation.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create the destination workbook (empty)
                Workbook mergedWorkbook = new Workbook();

                // Paths to source workbooks
                string sourcePath1 = "source1.xlsx";
                string sourcePath2 = "source2.xlsx";

                // Verify source files exist
                if (!File.Exists(sourcePath1))
                {
                    Console.WriteLine($"File not found: {sourcePath1}");
                    return;
                }

                if (!File.Exists(sourcePath2))
                {
                    Console.WriteLine($"File not found: {sourcePath2}");
                    return;
                }

                // Load source workbooks
                Workbook source1 = new Workbook(sourcePath1);
                Workbook source2 = new Workbook(sourcePath2);

                // Helper to copy all worksheets from a source workbook into the merged workbook
                void CopyWorksheets(Workbook source)
                {
                    foreach (Worksheet srcWs in source.Worksheets)
                    {
                        try
                        {
                            // Add a new empty worksheet to the merged workbook
                            int newIndex = mergedWorkbook.Worksheets.Add();
                            Worksheet destWs = mergedWorkbook.Worksheets[newIndex];

                            // Copy the source worksheet into the newly added worksheet
                            srcWs.Copy(destWs);
                        }
                        catch (Exception exCopy)
                        {
                            Console.WriteLine($"Error copying worksheet '{srcWs.Name}': {exCopy.Message}");
                        }
                    }
                }

                // Merge worksheets from both sources
                CopyWorksheets(source1);
                CopyWorksheets(source2);

                // Optional: remove the default empty sheet if it still exists and has no data
                if (mergedWorkbook.Worksheets.Count > 0 && mergedWorkbook.Worksheets[0].Cells.MaxDataColumn == -1 && mergedWorkbook.Worksheets[0].Cells.MaxDataRow == -1)
                {
                    mergedWorkbook.Worksheets.RemoveAt(0);
                }

                // Calculate all formulas in the merged workbook
                mergedWorkbook.CalculateFormula();

                // Flag to indicate presence of #REF! errors
                bool hasRefError = false;

                // Scan every cell in every worksheet for #REF! errors
                foreach (Worksheet ws in mergedWorkbook.Worksheets)
                {
                    Cells cells = ws.Cells;
                    foreach (Cell cell in cells)
                    {
                        try
                        {
                            // If the cell is an error and its string representation is #REF!
                            if (cell.Type == CellValueType.IsError && cell.StringValue == "#REF!")
                            {
                                hasRefError = true;
                                Console.WriteLine($"#REF! error found in sheet '{ws.Name}' at cell {cell.Name}");
                            }
                        }
                        catch (Exception exCell)
                        {
                            Console.WriteLine($"Error processing cell {cell.Name} in sheet '{ws.Name}': {exCell.Message}");
                        }
                    }
                }

                if (!hasRefError)
                {
                    Console.WriteLine("No #REF! errors detected after merging.");
                }

                // Save the merged workbook
                string outputPath = "merged_output.xlsx";
                mergedWorkbook.Save(outputPath);
                Console.WriteLine($"Merged workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
