// Title: Sync Excel worksheets by copying only changed cells with Aspose.Cells for .NET (C#)
// Description: Loads a source and a target workbook, determines the combined used range, iterates each cell, compares values (including null handling), copies differing values while preserving data types, and saves the updated workbook. Ideal for incremental Excel updates using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | .NET | worksheet synchronization | copy changed cells | compare Excel cells | incremental workbook update | Excel cell iteration | preserve data types | Excel merge C#
// Common Searches: Aspose.Cells sync two worksheets C# | Copy only modified cells between Excel files using Aspose | How to compare and update Excel cells with Aspose.Cells | Incremental Excel workbook merge .NET | Iterate over cells and copy differences Aspose.Cells
// Developer Intent: Update a target worksheet by copying only those cells whose values differ from a source worksheet, then save the result.
// Use Cases: Refresh a master report with daily changes from a temporary workbook. | Merge template edits into an existing data file without overwriting unchanged cells. | Create a version‑controlled copy of a workbook after applying selective updates. | Synchronize regional Excel sheets into a central repository while preserving original data.
// AI Prompts: Write C# code using Aspose.Cells that copies only changed cell values from one worksheet to another. | Explain best practices for handling null cells and preserving data types during Excel synchronization with Aspose.Cells. | Optimize the cell‑iteration loop to skip empty rows and reduce memory consumption in a worksheet sync routine.

using System;
using Aspose.Cells;

namespace AsposeCellsSyncDemo
{
    // Loads a source and a target workbook, determines the combined used range, iterates each cell, compares values (including null handling), copies differing values while preserving data types, and saves the updated workbook. Ideal for incremental Excel updates using Aspose.Cells in C#.
    public class WorksheetSynchronizer
    {
        /// <param name="sourceFile">Path to the source workbook.</param>
        /// <param name="targetFile">Path to the target workbook (will be updated).</param>
        /// <param name="outputFile">Path where the synchronized workbook will be saved.</param>
        public static void SyncWorksheets(string sourceFile, string targetFile, string outputFile)
        {
            // Load the source and target workbooks (creation and loading are done via Aspose.Cells APIs)
            Workbook sourceWorkbook = new Workbook(sourceFile);
            Workbook targetWorkbook = new Workbook(targetFile);

            // Assume we are working with the first worksheet of each workbook.
            // Adjust the index or name as needed.
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Worksheet targetSheet = targetWorkbook.Worksheets[0];

            // Get the Cells collections for easier access.
            Cells sourceCells = sourceSheet.Cells;
            Cells targetCells = targetSheet.Cells;

            // Determine the range to iterate: the maximum of source and target used rows/columns.
            int maxRow = Math.Max(sourceCells.MaxDataRow, targetCells.MaxDataRow);
            int maxColumn = Math.Max(sourceCells.MaxDataColumn, targetCells.MaxDataColumn);

            // Iterate through each cell within the determined range.
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    // Retrieve the source and target cells (may be null if the cell does not exist).
                    Cell sourceCell = sourceCells[row, col];
                    Cell targetCell = targetCells[row, col];

                    // If the source cell is null, there is nothing to copy.
                    if (sourceCell == null)
                        continue;

                    // Compare the values. Use the .Value property which returns the underlying object.
                    // Null handling is required because targetCell may be null.
                    object sourceValue = sourceCell.Value;
                    object targetValue = targetCell?.Value;

                    // If values are different (including one being null), copy the source value to target.
                    bool valuesDiffer = (sourceValue == null && targetValue != null) ||
                                        (sourceValue != null && !sourceValue.Equals(targetValue));

                    if (valuesDiffer)
                    {
                        // Ensure the target cell exists before putting a value.
                        if (targetCell == null)
                        {
                            targetCell = targetCells[row, col];
                        }

                        // Copy the value (preserves data type).
                        targetCell.PutValue(sourceValue);
                    }
                }
            }

            // Save the updated target workbook to the specified output file.
            targetWorkbook.Save(outputFile);
        }

        // Example usage
        public static void Main()
        {
            string sourcePath = "SourceWorkbook.xlsx";
            string targetPath = "TargetWorkbook.xlsx";
            string outputPath = "SynchronizedWorkbook.xlsx";

            SyncWorksheets(sourcePath, targetPath, outputPath);

            Console.WriteLine("Synchronization complete. Output saved to: " + outputPath);
        }
    }
}
