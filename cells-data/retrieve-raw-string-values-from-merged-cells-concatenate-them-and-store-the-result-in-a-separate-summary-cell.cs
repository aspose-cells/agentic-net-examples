using System;
using Aspose.Cells;

namespace MergedCellsConcatenationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ----- Sample data: create some merged cells with string values -----
            // Merge A1:B1 and put a value
            cells.Merge(0, 0, 1, 2);
            cells[0, 0].PutValue("First Part");

            // Merge C2:D3 (2 rows x 2 columns) and put a value
            cells.Merge(1, 2, 2, 2);
            cells[1, 2].PutValue("Second Part");

            // Merge E5 (single cell, not merged) – should be ignored in merged processing
            cells[4, 4].PutValue("Standalone");

            // ----- Retrieve raw string values from all merged cells -----
            // Get all merged areas in the worksheet
            CellArea[] mergedAreas = cells.GetMergedAreas();

            // StringBuilder for efficient concatenation
            System.Text.StringBuilder concatenated = new System.Text.StringBuilder();

            foreach (CellArea area in mergedAreas)
            {
                // Iterate through each cell in the merged area
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        // Get the raw string value (unformatted) of the cell
                        string rawValue = cells[row, col].StringValue ?? string.Empty;

                        // Append the value if it's not empty
                        if (!string.IsNullOrEmpty(rawValue))
                        {
                            if (concatenated.Length > 0)
                                concatenated.Append(" "); // separator between values
                            concatenated.Append(rawValue);
                        }
                    }
                }
            }

            // ----- Store the concatenated result in a summary cell (e.g., G1) -----
            cells["G1"].PutValue(concatenated.ToString());

            // Save the workbook
            workbook.Save("MergedCellsSummary.xlsx");
        }
    }
}