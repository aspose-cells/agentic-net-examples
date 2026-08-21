// Title: Aspose.Cells .NET – Detect merged cells and handle only the top‑left cell of each range
// Description: Creates a workbook, merges sample ranges, then iterates the used cells. By checking Cell.IsMerged and using GetMergedRange, the code processes only the first cell (FirstRow/FirstColumn) of each merged area, logs its address and value, and saves the file.
// Keywords: Aspose.Cells detect merged cells | C# merged range top left cell | Cell.IsMerged property | GetMergedRange example | enumerate used range Aspose.Cells | skip duplicate merged cells | Aspose.Cells .NET tutorial | global Aspose.Cells guide
// Common Searches: how to find merged cells in Aspose.Cells C# | process only the first cell of a merged range Aspose.Cells | skip duplicate cells in merged areas Aspose.Cells .NET | enumerate worksheet cells and detect merges Aspose | Aspose.Cells GetMergedRange top‑left cell
// Developer Intent: Identify merged cells while looping through a worksheet and execute logic only for the primary (top‑left) cell of each merged block.
// Use Cases: Log or export the address and value of each merged region’s leading cell. | Apply formatting, formulas, or calculations exclusively to the top‑left cell of merged areas. | Generate clean CSV/JSON output where merged cells are represented by their primary cell value.
// AI Prompts: Write C# code with Aspose.Cells that iterates a worksheet, detects merged cells, and runs an action only on the top‑left cell of each merged range. | Show how to use Cell.IsMerged and GetMergedRange to filter out duplicate cells in merged areas in Aspose.Cells for .NET. | Explain an efficient pattern for processing merged cells without redundant iterations in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook, merges sample ranges, then iterates the used cells. By checking Cell.IsMerged and using GetMergedRange, the code processes only the first cell (FirstRow/FirstColumn) of each merged area, logs its address and value, and saves the file.
class DetectMergedCells
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge some sample ranges
            cells.Merge(0, 0, 2, 2); // A1:B2
            cells[0, 0].PutValue("Merged A1");
            cells.Merge(2, 3, 3, 2); // D3:E5
            cells[2, 3].PutValue("Merged D3");

            // Determine the used range limits
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Enumerate all cells in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    if (cell.IsMerged)
                    {
                        // Retrieve the merged range that this cell belongs to
                        AsposeRange mergedRange = cell.GetMergedRange();

                        // Process only the top‑left cell of the merged area
                        if (mergedRange != null &&
                            mergedRange.FirstRow == row &&
                            mergedRange.FirstColumn == col)
                        {
                            // Example processing: output address and value
                            Console.WriteLine($"Top‑left merged cell: {cell.Name}, Value: {cell.StringValue}");
                        }
                    }
                    else
                    {
                        // Example processing for non‑merged cells
                        Console.WriteLine($"Normal cell: {cell.Name}, Value: {cell.StringValue}");
                    }
                }
            }

            // Save the workbook
            string outputPath = "MergedCellsProcessed.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
