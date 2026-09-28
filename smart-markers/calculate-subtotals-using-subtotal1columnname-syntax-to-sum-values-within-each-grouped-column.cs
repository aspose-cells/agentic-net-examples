// Title: Add grouped SUM subtotals to an Excel sheet with Aspose.Cells Subtotal method (subtotal1:ColumnName) in C#
// AI Prompts: Generate a C# program that fills an Excel worksheet with Category, Item, and Amount columns and uses the Cells.Subtotal method to insert SUM subtotals for the Amount column grouped by Category via the (subtotal1:ColumnName) syntax. | Extend the example to add a Quantity column and switch the subtotal function to Average for both Amount and Quantity while keeping page breaks and placing the summary row below the data. | Write C# code that applies multiple subtotal functions (e.g., SUM for Amount, MAX for Discount) to each Category group, customizing the position of the subtotal rows with the Subtotal method.
// Common Searches: aspnet cells subtotal method group by column c# example | how to add sum subtotals per category in Excel using Aspose.Cells C# | using (subtotal1:ColumnName) syntax with Aspose.Cells Subtotal function | Aspose.Cells create subtotal rows with page breaks in C# workbook | C# calculate grouped subtotals in Excel file with Aspose.Cells
// Tags: Subtotal API in Aspose.Cells C# | grouped column sum with ConsolidationFunction.Sum | subtotal1 column syntax for Excel automation | insert page breaks using Subtotal function | category based subtotals in workbook

using Aspose.Cells;
using System;

// The sample creates a new workbook, writes Category, Item, and Amount data, defines the range A1:C6, and calls Cells.Subtotal to group rows by the Category column (index 0) and insert SUM subtotals for the Amount column (index 2). Existing subtotals are replaced, page breaks are added, and the summary row is placed below the data before saving as SubtotalResult.xlsx.
class SubtotalExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data with a header row
        // Columns: Category (A), Item (B), Amount (C)
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Item");
        cells["C1"].PutValue("Amount");

        object[,] data = new object[,]
        {
            { "A", "Item1", 100 },
            { "A", "Item2", 150 },
            { "B", "Item3", 200 },
            { "B", "Item4", 250 },
            { "A", "Item5", 120 }
        };

        // Fill the worksheet with the sample data starting from row 2 (zero‑based index 1)
        for (int i = 0; i < data.GetLength(0); i++)
        {
            cells[i + 1, 0].PutValue(data[i, 0]); // Category
            cells[i + 1, 1].PutValue(data[i, 1]); // Item
            cells[i + 1, 2].PutValue(data[i, 2]); // Amount
        }

        // Define the cell area that includes the header and all data rows
        CellArea area = CellArea.CreateCellArea("A1", "C6");

        // Apply subtotals:
        // - Group by the first column (Category) -> index 0
        // - Use SUM function
        // - Add subtotal to the third column (Amount) -> index 2
        // - Replace existing subtotals, insert page breaks, place summary below data
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 }, true, true, true);

        // Save the workbook with the applied subtotals
        workbook.Save("SubtotalResult.xlsx");
    }
}
