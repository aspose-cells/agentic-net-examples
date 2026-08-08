// Title: Aspose.Cells .NET: Concatenate Raw Strings from All Merged Cells into a Summary Cell
// Description: C# code that builds a workbook, creates merged ranges, extracts each cell's raw string value (without formatting), joins non‑empty texts with a StringBuilder, writes the combined result to a chosen summary cell (e.g., G1), and saves the workbook.
// Keywords: Aspose.Cells merged cells | C# get merged cell values | concatenate merged cell text | summary cell Aspose.Cells | StringBuilder Excel .NET | raw string value merged range | Excel merged cells API | Aspose.Cells .NET example | C# Excel automation | US developers Aspose.Cells
// Common Searches: How to read raw values from merged cells using Aspose.Cells C# | Combine text from multiple merged ranges into one cell Aspose.Cells | Aspose.Cells .NET concatenate merged cell strings | Extract unformatted strings from merged cells Aspose | Save merged cell summary in Excel with Aspose.Cells
// Developer Intent: Read every merged area, pull each cell's raw text, merge the non‑empty strings, and place the final string into a designated summary cell.
// Use Cases: Create an index cell that aggregates titles stored in merged header regions. | Summarize notes entered across several merged comment blocks before exporting. | Generate a dashboard label by joining labels from multiple merged sections into a single cell.
// AI Prompts: Provide a C# Aspose.Cells snippet that iterates all merged areas, collects raw string values, concatenates them with a space, and writes the result to cell G1. | Show how to use StringBuilder with Aspose.Cells to build a summary string from non‑empty merged cells in a worksheet. | Explain the steps to retrieve unformatted text from merged cells, handle empty cells, and store the combined output in a separate summary cell.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsMergedCellSummary
{
    // C# code that builds a workbook, creates merged ranges, extracts each cell's raw string value (without formatting), joins non‑empty texts with a StringBuilder, writes the combined result to a chosen summary cell (e.g., G1), and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Sample data: create a few merged ranges and put values in them
            // ------------------------------------------------------------
            // Merge A1:B2 and set a value
            cells.Merge(0, 0, 2, 2);               // A1:B2
            cells[0, 0].PutValue("First");

            // Merge C3:D4 and set a value
            cells.Merge(2, 2, 2, 2);               // C3:D4
            cells[2, 2].PutValue("Second");

            // Merge E5:E6 (single column) and set a value
            cells.Merge(4, 4, 2, 1);               // E5:E6
            cells[4, 4].PutValue("Third");

            // ------------------------------------------------------------
            // Retrieve raw string values from all merged cells,
            // concatenate them, and store the result in a summary cell.
            // ------------------------------------------------------------
            // Get all merged areas in the worksheet
            CellArea[] mergedAreas = cells.GetMergedAreas();

            // Use StringBuilder for efficient concatenation
            StringBuilder summaryBuilder = new StringBuilder();

            foreach (CellArea area in mergedAreas)
            {
                // Iterate through each cell inside the merged area
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        // Retrieve the raw string value (no formatting)
                        string rawValue = cells[row, col].StringValue ?? string.Empty;

                        // Append the value if it's not empty
                        if (!string.IsNullOrEmpty(rawValue))
                        {
                            // Separate values with a space (customize as needed)
                            if (summaryBuilder.Length > 0)
                                summaryBuilder.Append(' ');
                            summaryBuilder.Append(rawValue);
                        }
                    }
                }
            }

            // Write the concatenated result to a separate summary cell (e.g., G1)
            cells["G1"].PutValue(summaryBuilder.ToString());

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("MergedCellsSummary.xlsx");
        }
    }
}
