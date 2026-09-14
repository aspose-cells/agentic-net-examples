// Title: Add Max subtotals to column C with group summary rows placed after each category using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that groups rows by the first column and adds a row after each group showing the highest value from column C using Aspose.Cells. | Show how to invoke the subtotal method with the max consolidation function, keep existing data, and suppress page breaks while placing summary rows below each category.
// Common Searches: Aspose.Cells example for inserting a row with the highest value per category in C# | How to calculate maximums per group without page breaks using Aspose.Cells subtotal feature in .NET | C# code to generate an Excel file with subtotal rows placed after each group using Aspose.Cells | Create a grouped Excel report with max values per group using Aspose.Cells for .NET
// Tags: max value subtotal Aspose.Cells | grouped summary rows below each category | suppress page breaks in subtotal operation | subtotal on column C using max consolidation | excel workbook generation C# Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    // The example creates a workbook, fills it with sample data containing a Category column, then uses Aspose.Cells to add a subtotal row that shows the maximum value from column C for each category. The summary rows are inserted directly below each group, page breaks are disabled, and the workbook is saved as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data
            // Header row
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Value");

            // Data rows
            object[,] data = new object[,]
            {
                { "Group1", "A", 10 },
                { "Group1", "B", 20 },
                { "Group1", "C", 15 },
                { "Group2", "A", 30 },
                { "Group2", "B", 25 },
                { "Group2", "C", 35 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                for (int c = 0; c < data.GetLength(1); c++)
                {
                    cells[r + 1, c].PutValue(data[r, c]);
                }
            }

            // Define the range that contains the data (including header)
            // A1:C7  -> rows 0-6, columns 0-2
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = data.GetLength(0), // includes header row
                EndColumn = 2
            };

            // Add subtotals:
            // - Group by first column (Category) -> index 0
            // - Use Max function on column C (index 2)
            // - Do not replace existing subtotals, no page breaks, summary rows below data
            cells.Subtotal(
                area,
                0,                                 // groupBy column index
                ConsolidationFunction.Max,         // Max function
                new int[] { 2 },                   // subtotal on column C
                false,                             // replace existing subtotals
                false,                             // page breaks between groups
                true                               // summary rows placed below each group
            );

            // Save the workbook
            workbook.Save("Subtotal_Max_ColumnC_AfterEachGroup.xlsx");
        }
    }
}
