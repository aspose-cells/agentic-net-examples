// Title: Enumerate merged cells and retrieve their ranges in an Aspose.Cells workbook with C#
// AI Prompts: Write C# code using Aspose.Cells to loop through all used cells in a worksheet, detect cells where IsMerged is true, and output each cell's address together with its merged range. | Show how to invoke the GetMergedRange method on a merged cell in Aspose.Cells .NET and extract the range reference for further processing. | Provide an example that saves the workbook after handling merged cells, including proper exception handling and resource cleanup.
// Common Searches: C# Aspose.Cells find all merged cells in a worksheet | How to get merged range address for a cell using Aspose.Cells .NET | Iterate over used cells and check IsMerged property in Aspose.Cells | Save Excel file after processing merged cells with Aspose.Cells C#
// Tags: Aspose.Cells C# merged cell enumeration | merged range extraction Aspose.Cells | IsMerged property usage Aspose.Cells | processing merged ranges Aspose.Cells | save workbook after merged cell handling

using Aspose.Cells;
using System;

// The sample creates a workbook, merges two ranges, adds values, then iterates over the used cells. For each cell it checks the IsMerged flag, obtains the merged range via GetMergedRange, prints the cell address and range, and finally saves the file as MergedCellsProcessed.xlsx.
class IdentifyMergedCells
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
            cells.Merge(3, 2, 3, 2); // C4:D6

            // Add values to demonstrate merged and normal cells
            cells["A1"].PutValue("Merged A1");
            cells["C4"].PutValue("Merged C4");
            cells["E1"].PutValue("Normal");

            // Determine the used range of the sheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Enumerate all cells and process those that are part of a merged range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.IsMerged)
                    {
                        // Retrieve the merged range that this cell belongs to
                        Aspose.Cells.Range mergedRange = cell.GetMergedRange();

                        // Example processing: output merged cell information
                        Console.WriteLine($"Cell {cell.Name} is merged. Merged range: {mergedRange.RefersTo}");
                    }
                }
            }

            // Save the workbook
            workbook.Save("MergedCellsProcessed.xlsx");
            Console.WriteLine("Workbook saved as MergedCellsProcessed.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
