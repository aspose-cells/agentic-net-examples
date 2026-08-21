// Title: Add Min Subtotal to Column G and Show Flat Summary Rows with Aspose.Cells for .NET
// Description: Creates a workbook, populates columns A‑G, defines range A1:G11, and uses Worksheet.Cells.Subtotal with ConsolidationFunction.Min to insert minimum subtotals for column G. The summary is placed below the data, existing subtotals are replaced, page breaks are suppressed, and outline grouping is turned off (SummaryRowBelow = false) for a flat view. The file is saved as Subtotal_Min_ColumnG_FlatView.xlsx.
// Keywords: Aspose.Cells | .NET | subtotal min | column G | disable outline | flat view | Worksheet.Cells.Subtotal | ConsolidationFunction.Min | summary row below | Excel automation
// Common Searches: Aspose.Cells add min subtotal column | disable outline grouping Aspose.Cells .NET | Worksheet.Cells.Subtotal example C# | flat view subtotal rows Aspose.Cells | C# Aspose.Cells subtotal without grouping
// Developer Intent: Generate an Excel workbook that calculates the minimum subtotal for column G and displays the summary rows without collapsible outlines.
// Use Cases: Financial reporting where the lowest transaction amount per category is highlighted as a subtotal in a flat list. | Inventory sheets that need minimum quantity subtotals per group while keeping the worksheet free of outline hierarchies. | Automated data exports that require min subtotals on a numeric column and a non‑collapsible presentation for downstream processing.
// AI Prompts: Write C# code using Aspose.Cells to add a Min subtotal on column G for a given range and disable outline grouping so the summary rows appear in a flat view. | Show an example of Worksheet.Cells.Subtotal with ConsolidationFunction.Min and set Worksheet.Outline.SummaryRowBelow to false in Aspose.Cells for .NET. | Explain how to replace existing subtotals, avoid page breaks, place the summary below the data, and turn off outline grouping when using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, populates columns A‑G, defines range A1:G11, and uses Worksheet.Cells.Subtotal with ConsolidationFunction.Min to insert minimum subtotals for column G. The summary is placed below the data, existing subtotals are replaced, page breaks are suppressed, and outline grouping is turned off (SummaryRowBelow = false) for a flat view. The file is saved as Subtotal_Min_ColumnG_FlatView.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (columns A to G). Column G (index 6) will hold numeric values.
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["G1"].PutValue("Amount");

        for (int i = 0; i < 10; i++)
        {
            // Example grouping column (A) – alternating groups
            worksheet.Cells[i + 1, 0].PutValue(i % 2 == 0 ? "Group1" : "Group2");
            // Column G values
            worksheet.Cells[i + 1, 6].PutValue((i + 1) * 10);
        }

        // Define the cell area that contains the data (A1:G11)
        CellArea area = CellArea.CreateCellArea("A1", "G11");

        // Add subtotals:
        // - Group by column G (index 6)
        // - Use the Min function
        // - Apply the subtotal to column G (index 6)
        // - Replace existing subtotals, no page breaks, summary placed below data
        worksheet.Cells.Subtotal(
            area,
            6,                                 // groupBy column (G)
            ConsolidationFunction.Min,         // Min function
            new int[] { 6 },                   // subtotal column (G)
            true,                              // replace existing subtotals
            false,                             // no page breaks between groups
            true                               // summary below data
        );

        // Disable outline grouping for a flat view (summary rows will not be shown as outline groups)
        worksheet.Outline.SummaryRowBelow = false;

        // Save the workbook
        workbook.Save("Subtotal_Min_ColumnG_FlatView.xlsx");
    }
}
