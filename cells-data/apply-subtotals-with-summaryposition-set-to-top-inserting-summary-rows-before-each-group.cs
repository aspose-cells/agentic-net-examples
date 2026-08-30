// Title: Add summary rows above each category group with SUM using Aspose.Cells Subtotal in C#
// AI Prompts: Generate C# code that uses Aspose.Cells to insert a summary row before each category group, calculating the sum of the Amount column. | Show how to set the Outline.SummaryRowBelow property so that subtotal rows appear at the top of each group in an Aspose.Cells workbook. | Write a reusable method that takes a worksheet, a data range, and a list of numeric columns, then adds top‑positioned subtotal rows via the Subtotal API.
// Common Searches: Aspose.Cells C# how to place subtotal rows at the top of grouped data | C# example for adding sum subtotals before each category in an Excel file with Aspose.Cells | Using Subtotal method to create top summary rows for a column range in Aspose.Cells .NET
// Tags: Aspose.Cells Subtotal top summary rows | C# insert subtotal before group | Excel workbook sum subtotals with Aspose.Cells | outline summary row placement Aspose.Cells | group by column subtotal C# Aspose.Cells

using Aspose.Cells;
using System;

// // Demonstrates creating a workbook, populating Category and Amount columns, and using Aspose.Cells.Subtotal together with Outline.SummaryRowBelow = false to insert summary rows above each category group using the SUM function.
class SubtotalTopDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (Category and Amount)
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("A");
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["A4"].PutValue("B");
        worksheet.Cells["B4"].PutValue(200);
        worksheet.Cells["A5"].PutValue("B");
        worksheet.Cells["B5"].PutValue(250);

        // Define the range that contains the data
        CellArea dataArea = CellArea.CreateCellArea("A1", "B5");

        // Apply subtotals:
        // - Group by the first column (Category) -> groupBy = 0
        // - Use SUM function for subtotals
        // - Subtotal the second column (Amount) -> totalList = new int[] { 1 }
        // - Replace existing subtotals = true
        // - Do not insert page breaks between groups = false
        // - Place summary rows above the detail rows (top) = false
        worksheet.Cells.Subtotal(
            dataArea,
            0,
            ConsolidationFunction.Sum,
            new int[] { 1 },
            true,
            false,
            false);

        // Ensure the outline setting also places summary rows above the data
        worksheet.Outline.SummaryRowBelow = false;

        // Save the workbook
        workbook.Save("SubtotalTopDemo.xlsx");
    }
}
