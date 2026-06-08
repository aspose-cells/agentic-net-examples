using System;
using Aspose.Cells;

namespace AsposeCellsMergedCellDetection
{
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

                // Sample data: create some merged cells and put values
                cells.Merge(0, 0, 2, 2); // A1:C3 (rows 0‑2, columns 0‑2)
                cells[0, 0].PutValue("Merged A1:C3");

                cells.Merge(3, 1, 5, 3); // B4:D6 (rows 3‑5, columns 1‑3)
                cells[3, 1].PutValue("Merged B4:D6");

                cells[6, 0].PutValue("Normal Cell A7");

                // Enumerate all cells in the used range
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Check if the cell is part of a merged range
                        if (cell.IsMerged)
                        {
                            // Get the merged range that this cell belongs to
                            Aspose.Cells.Range mergedRange = cell.GetMergedRange();

                            // Process only the top‑left cell of the merged area
                            if (mergedRange != null &&
                                row == mergedRange.FirstRow &&
                                col == mergedRange.FirstColumn)
                            {
                                Console.WriteLine($"Top‑left merged cell {cell.Name} value: {cell.StringValue}");
                                // Custom processing for merged cells can be added here
                            }
                        }
                        else
                        {
                            // Process non‑merged cells as needed
                            Console.WriteLine($"Normal cell {cell.Name} value: {cell.StringValue}");
                        }
                    }
                }

                // Save the workbook
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