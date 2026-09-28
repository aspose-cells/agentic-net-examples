// Title: Calculate average subtotals for column F and place summary rows above data using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that groups rows by the first column, computes the average of column F, and inserts the subtotal rows at the top of the worksheet with Aspose.Cells. | Show how to call the Cells.Subtotal method with summaryBelowData set to false to generate top‑positioned average summaries in an Excel file.
// Common Searches: Aspose.Cells C# subtotal average of column F with summary rows on top | How to add average subtotal rows above data in Excel using Aspose.Cells .NET | C# Aspose.Cells Subtotal method summaryBelowData false example | Group by Category column and calculate average Score subtotal with Aspose.Cells | Insert top summary rows for average calculation in worksheet using Aspose.Cells
// Tags: Aspose.Cells subtotal average column F | disable summaryBelowData Aspose.Cells | C# group rows by column subtotal | insert top summary rows Excel .NET | average aggregation with Aspose.Cells Subtotal

using System;
using Aspose.Cells;

namespace SubtotalAverageTopDemo
{
    // // Demonstrates creating a workbook, populating it with sample data, and using Aspose.Cells' Cells.Subtotal method to group by the Category column, calculate the average of the Score column (F), insert the subtotal rows at the top of the sheet, and save the file as SubtotalAverageTopDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (including column F which is index 5)
            // Header row
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Qty");
            cells["D1"].PutValue("Price");
            cells["E1"].PutValue("Discount");
            cells["F1"].PutValue("Score"); // Column F to average

            // Sample data rows
            object[,] data = new object[,]
            {
                {"A","Item1",10,5.0,0.1,80},
                {"A","Item2",15,7.5,0.05,85},
                {"B","Item3",8,12.0,0.2,78},
                {"B","Item4",20,3.5,0.0,90},
                {"C","Item5",5,20.0,0.15,88}
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    cells[i + 1, j].PutValue(data[i, j]);
                }
            }

            // Define the cell area that contains the data (A1:F6)
            CellArea area = CellArea.CreateCellArea("A1", "F6");

            // Apply subtotal:
            // - Group by the first column (Category) -> index 0
            // - Use Average function on column F (index 5)
            // - Do not replace existing subtotals, no page breaks, summary rows at the top (summaryBelowData = false)
            cells.Subtotal(
                area,
                0,                                 // groupBy column index
                ConsolidationFunction.Average,    // average function
                new int[] { 5 },                  // apply to column F
                false,                            // replace existing subtotals
                false,                            // page breaks between groups
                false);                           // summaryBelowData = false (summary at top)

            // Save the workbook
            workbook.Save("SubtotalAverageTopDemo.xlsx");
        }
    }
}
