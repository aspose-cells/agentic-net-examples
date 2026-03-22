using System;
using Aspose.Cells;

namespace CombineWorkbooksExample
{
    class Program
    {
        static void Main()
        {
            // Paths of source workbooks to be combined
            string[] sourceFiles = new string[]
            {
                "Source1.xlsx",
                "Source2.xlsx"
                // Add more file paths as needed
            };

            // Create the destination workbook (empty)
            Workbook destWorkbook = new Workbook();

            // Combine each source workbook into the destination workbook
            foreach (string file in sourceFiles)
            {
                // Load source workbook
                Workbook srcWorkbook = new Workbook(file);

                // Merge the source workbook into the destination workbook
                // This adds all worksheets from the source workbook
                destWorkbook.Combine(srcWorkbook);
            }

            // Add a new worksheet that will hold the aggregated data
            Worksheet aggSheet = destWorkbook.Worksheets.Add("Aggregated");

            // Row index in the aggregated sheet where the next source data will be placed
            int aggCurrentRow = 0;

            // Iterate over all worksheets that were added by Combine (skip the aggregation sheet itself)
            foreach (Worksheet ws in destWorkbook.Worksheets)
            {
                if (ws.Name == "Aggregated")
                    continue; // Skip the aggregation sheet

                // Determine the used range of the current source worksheet
                int maxRow = ws.Cells.MaxDataRow;
                int maxCol = ws.Cells.MaxDataColumn;

                // Copy each non‑empty cell to the aggregation sheet
                for (int r = 0; r <= maxRow; r++)
                {
                    for (int c = 0; c <= maxCol; c++)
                    {
                        Cell srcCell = ws.Cells[r, c];
                        if (srcCell.Value != null)
                        {
                            // Preserve the original value type
                            aggSheet.Cells[aggCurrentRow + r, c].PutValue(srcCell.Value);
                        }
                    }
                }

                // Add a blank row between data blocks for readability
                aggCurrentRow += maxRow + 2;
            }

            // Save the combined workbook with the aggregated sheet
            destWorkbook.Save("CombinedWithAggregatedSheet.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbooks combined and data aggregated successfully.");
        }
    }
}