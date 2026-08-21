// Title: C# – Enumerate cells and handle only the top‑left cell of each merged range using IsMerged & GetMergedRange (Aspose.Cells)
// Description: Creates a workbook, merges two areas, iterates through all populated cells, checks Cell.IsMerged, obtains the corresponding Range with GetMergedRange, and processes only the first cell of each merged block to prevent duplicate handling, then saves the file.
// Keywords: Aspose.Cells C# merged cells | Cell.IsMerged | GetMergedRange | enumerate worksheet cells | merged range detection | Aspose.Cells .NET example | process merged areas | top left merged cell
// Common Searches: Aspose.Cells find merged cells C# | How to use IsMerged property in Aspose.Cells | Get merged range from a cell Aspose.Cells .NET | Iterate worksheet cells without duplicate merged processing | C# Aspose.Cells top left cell of merged area
// Developer Intent: Detect cells that belong to a merged block and execute logic once per merged region.
// Use Cases: Log the address and size of each merged block for debugging or reporting. | Apply a specific style (e.g., background color) to the leading cell of every merged area. | Extract header values from merged cells to build a dictionary for downstream data processing.
// AI Prompts: Write a method that returns a collection of all merged Range objects in a given worksheet using Aspose.Cells. | Modify the example to copy the value from the top‑left cell of each merged range into a new worksheet column. | Generate code that unmerges all merged ranges after they have been processed.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, merges two areas, iterates through all populated cells, checks Cell.IsMerged, obtains the corresponding Range with GetMergedRange, and processes only the first cell of each merged block to prevent duplicate handling, then saves the file.
class ProcessMergedCells
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create some merged ranges for demonstration
            cells.Merge(0, 0, 2, 2); // Merge A1:B2
            cells.Merge(3, 1, 3, 3); // Merge B4:D6

            // Put values in the merged cells (optional)
            cells["A1"].PutValue("Header1");
            cells["B4"].PutValue("Header2");

            // Determine the area that contains data
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Enumerate all cells and process those that belong to merged ranges
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the current cell is part of a merged range
                    if (cell.IsMerged)
                    {
                        // Retrieve the merged range for the cell
                        AsposeRange mergedRange = cell.GetMergedRange();

                        // Process only the top‑left cell of each merged area to avoid duplicates
                        if (cell.Row == mergedRange.FirstRow && cell.Column == mergedRange.FirstColumn)
                        {
                            Console.WriteLine($"Merged range starting at {cell.Name}: " +
                                              $"Rows={mergedRange.RowCount}, Columns={mergedRange.ColumnCount}");
                            // Additional processing can be placed here (e.g., reading/writing values)
                        }
                    }
                }
            }

            // Save the workbook
            workbook.Save("ProcessedMergedCells.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
