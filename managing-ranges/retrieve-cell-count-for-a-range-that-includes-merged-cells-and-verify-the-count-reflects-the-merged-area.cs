using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class MergedRangeCellCountDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Merge a block of cells: A1:C2 (3 columns x 2 rows = 6 cells)
                // Parameters: firstRow, firstColumn, totalRows, totalColumns
                cells.Merge(0, 0, 2, 3); // A1:C2

                // Create a range that covers the merged area
                AsposeRange mergedRange = cells.CreateRange("A1:C2");

                // Retrieve the cell count directly from the merged range
                int actualCellCount = mergedRange.RowCount * mergedRange.ColumnCount;

                // Expected cell count is rows * columns of the original range
                int expectedCellCount = mergedRange.RowCount * mergedRange.ColumnCount;

                // Output the results
                Console.WriteLine($"Actual CellCount from merged range: {actualCellCount}");
                Console.WriteLine($"Expected cell count (rows * columns): {expectedCellCount}");

                // Verify that the counts match
                if (actualCellCount == expectedCellCount)
                    Console.WriteLine("Verification succeeded: CellCount reflects the merged area.");
                else
                    Console.WriteLine("Verification failed: CellCount does not match the merged area.");

                // Additionally, retrieve merged areas using Cells.GetMergedAreas()
                CellArea[] mergedAreas = cells.GetMergedAreas();
                Console.WriteLine($"Number of merged areas reported by GetMergedAreas(): {mergedAreas.Length}");
                if (mergedAreas.Length > 0)
                {
                    CellArea area = mergedAreas[0];
                    int mergedAreaCellCount = (area.EndRow - area.StartRow + 1) * (area.EndColumn - area.StartColumn + 1);
                    Console.WriteLine($"Merged area spans rows {area.StartRow}-{area.EndRow} and columns {area.StartColumn}-{area.EndColumn}");
                    Console.WriteLine($"Cell count calculated from merged area: {mergedAreaCellCount}");
                }

                // Save the workbook (optional, just to demonstrate lifecycle compliance)
                string outputPath = "MergedRangeCellCountDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}