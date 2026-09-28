// Title: How to add variance subtotals on column I with outline view using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that groups rows by column A, calculates variance subtotals for column I with Aspose.Cells, and enables outline view with summary rows positioned below the details. | Demonstrate defining a CellArea, applying ConsolidationFunction.Var via the Subtotal method, turning on sheet.Outline.SummaryRowBelow, and saving the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# subtotal variance function column I example | Enable outline summary rows below data in Aspose.Cells .NET | Group data by column and calculate variance subtotal using Aspose.Cells | How to use ConsolidationFunction.Var with Subtotal method in C# | Create hierarchical outline with variance subtotals in Aspose.Cells workbook
// Tags: subtotal variance Aspose.Cells C# | outline view summary rows below Aspose.Cells | group by column A subtotal var function | ConsolidationFunction.Var range application | save workbook as .xlsx Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalVarExample
{
    // The example creates a workbook, populates column A with group identifiers and column I with numeric values, defines a CellArea covering the data, adds variance subtotals on column I grouped by column A using ConsolidationFunction.Var, enables outline view with summary rows placed below the detail rows, and saves the result as SubtotalVarOutline.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (headers in row 1)
            // Columns A-H are grouping fields, column I (index 8) contains numeric values
            cells["A1"].PutValue("Group");
            cells["I1"].PutValue("Value");

            // Sample data rows
            object[,] data = new object[,]
            {
                { "A", 10 },
                { "A", 20 },
                { "A", 30 },
                { "B", 15 },
                { "B", 25 },
                { "B", 35 },
                { "C", 12 },
                { "C", 22 },
                { "C", 32 }
            };

            // Fill data starting from row 2 (zero‑based index 1)
            for (int r = 0; r < data.GetLength(0); r++)
            {
                // Column A (index 0) for group
                cells[r + 1, 0].PutValue(data[r, 0]);
                // Column I (index 8) for value
                cells[r + 1, 8].PutValue(data[r, 1]);
            }

            // Define the range that includes the header and all data rows
            // StartRow = 0, StartColumn = 0 (A), EndRow = data rows count, EndColumn = 8 (I)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = data.GetLength(0),
                EndColumn = 8
            };

            // Add subtotals:
            // - Group by column A (index 0)
            // - Use Var function (variance)
            // - Apply subtotal to column I (index 8)
            // - Replace existing subtotals, no page breaks, summary below data
            cells.Subtotal(area, 0, ConsolidationFunction.Var, new int[] { 8 }, true, false, true);

            // Enable outline view (summary rows positioned below detail rows)
            sheet.Outline.SummaryRowBelow = true;

            // Save the workbook
            workbook.Save("SubtotalVarOutline.xlsx");
        }
    }
}
