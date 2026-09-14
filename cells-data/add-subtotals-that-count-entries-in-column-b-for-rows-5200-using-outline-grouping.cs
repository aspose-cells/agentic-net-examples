// Title: Create a count subtotal for column B rows 5‑200 with outline grouping using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# program that defines a CellArea covering rows 5‑200 in column B, applies a Count aggregation as a subtotal, and enables an outline with grouped rows, placing the summary row after the detail rows using Aspose.Cells. | Add code to set Worksheet.Outline.SummaryRowBelow = true after creating the subtotal so the total appears beneath the grouped rows. | Save the workbook as an .xlsx file named SubtotalOutline.xlsx after configuring the subtotal and outline settings.
// Common Searches: Aspose.Cells how to add a count subtotal for a specific column range in C# | C# subtotal rows 5 to 200 column B with outline grouping Aspose.Cells | set summary row below detail rows after subtotal Aspose.Cells .NET | using range definition to apply subtotal function in Aspose.Cells example | outline grouping subtotal count non‑empty cells Aspose.Cells C#
// Tags: Aspose.Cells C# column B subtotal | outline hierarchy total row placement | range for rows 5 through 200 subtotal | ConsolidationFunction.Count example | Excel workbook subtotal outline C#

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    // The program creates a new workbook, defines a CellArea covering rows 5‑200 in column B, adds a Count subtotal for that range, configures the outline to place the summary row after the detail rows, and saves the file as SubtotalOutline.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Assume that column B (index 1) already contains data from
            // row 5 (zero‑based index 4) to row 200 (zero‑based index 199).
            // The following code adds a subtotal that counts the entries
            // in column B and creates an outline grouping.
            // ------------------------------------------------------------

            // Define the range that contains the data to be subtotaled.
            // Here we use only column B, but the range can be expanded if needed.
            CellArea area = new CellArea
            {
                StartRow = 4,      // Row 5 (zero‑based)
                EndRow = 199,      // Row 200 (zero‑based)
                StartColumn = 1,   // Column B (zero‑based)
                EndColumn = 1      // Column B
            };

            // Apply subtotal:
            // - groupBy = 0 because within the defined area the grouping column is the first (and only) column.
            // - ConsolidationFunction.Count counts the number of non‑empty cells.
            // - totalList = new int[] { 0 } adds the subtotal for the same column.
            cells.Subtotal(area, 0, ConsolidationFunction.Count, new int[] { 0 });

            // Configure outline to place the summary row below the detail rows.
            worksheet.Outline.SummaryRowBelow = true;

            // Save the workbook
            workbook.Save("SubtotalOutline.xlsx");
        }
    }
}
