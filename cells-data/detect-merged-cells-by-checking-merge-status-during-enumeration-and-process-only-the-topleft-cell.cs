// Title: Detect Top‑Left Cells of Merged Ranges While Enumerating Worksheets – Aspose.Cells for .NET
// Description: C# example that creates a workbook, merges A1:B2 and C3:D4, then loops through cells up to MaxDataRow/MaxDataColumn. It uses IsMerged and GetMergedRange to process only the first (top‑left) cell of each merged area and skips the rest, finally saving the file.
// Keywords: Aspose.Cells C# merged cells detection | GetMergedRange | IsMerged | top left merged cell | enumerate worksheet cells | skip inner merged cells | Excel merge detection .NET | Aspose.Cells sample code
// Common Searches: Aspose.Cells get first cell of merged range C# | how to ignore inner cells of a merged area in Aspose.Cells | enumerate worksheet data and skip duplicate merged cells | detect merged cells with GetMergedRange Aspose.Cells
// Developer Intent: Identify merged regions during worksheet iteration and handle only their top‑left cells.
// Use Cases: Export Excel data to CSV without duplicate values from merged cells. | Apply calculations or formatting exclusively to the leading cell of each merged block. | Generate a list of all merged ranges together with the value stored in their top‑left cell.
// AI Prompts: Write a C# method that returns a dictionary of top‑left merged cells and their values from a Worksheet using Aspose.Cells. | Provide code to copy only the top‑left cells of merged ranges from one worksheet to another, preserving formatting. | Explain how to modify the enumeration loop to skip empty rows while still detecting merged cells efficiently.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsMergedCellDetection
{
    // C# example that creates a workbook, merges A1:B2 and C3:D4, then loops through cells up to MaxDataRow/MaxDataColumn. It uses IsMerged and GetMergedRange to process only the first (top‑left) cell of each merged area and skips the rest, finally saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample merged ranges for demonstration
                // Merge A1:B2 (top‑left cell is A1)
                cells.Merge(0, 0, 2, 2);
                cells[0, 0].PutValue("Merged A1:B2");

                // Merge C3:D4 (top‑left cell is C3)
                cells.Merge(2, 2, 2, 2);
                cells[2, 2].PutValue("Merged C3:D4");

                // Put some normal (non‑merged) data
                cells[4, 0].PutValue("Normal Cell");

                // Enumerate all cells that contain data
                for (int row = 0; row <= cells.MaxDataRow; row++)
                {
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        Cell cell = cells[row, col];

                        // Skip empty cells
                        if (cell.Value == null) continue;

                        // Process only the top‑left cell of a merged area
                        if (cell.IsMerged)
                        {
                            // Get the merged range that this cell belongs to
                            AsposeRange mergedRange = cell.GetMergedRange();

                            // If the current cell is the first cell of the merged range, handle it
                            if (mergedRange != null &&
                                mergedRange.FirstRow == row &&
                                mergedRange.FirstColumn == col)
                            {
                                Console.WriteLine($"Top‑left merged cell {cell.Name}: {cell.Value}");
                                // Add custom processing logic here (e.g., export, modify, etc.)
                            }
                            // Otherwise, ignore the rest of the cells in the merged area
                        }
                        else
                        {
                            // Handle normal (non‑merged) cells
                            Console.WriteLine($"Normal cell {cell.Name}: {cell.Value}");
                            // Add custom processing logic here
                        }
                    }
                }

                // Save the workbook (demonstrates that the code compiles and runs)
                workbook.Save("MergedCellDetectionResult.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
