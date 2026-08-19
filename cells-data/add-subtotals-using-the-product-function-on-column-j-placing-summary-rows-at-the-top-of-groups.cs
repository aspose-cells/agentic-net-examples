// Title: C# – Add Product Subtotals Above Groups in Column J Using Aspose.Cells Subtotal
// Description: Creates a workbook, fills column I with group keys and column J with numeric values, defines a CellArea, and calls Cells.Subtotal to group by column I, calculate the Product of column J, and insert the subtotal rows above each group before saving as SubtotalProductDemo.xlsx.
// Keywords: Aspose.Cells Subtotal C# | ConsolidationFunction.Product | summary rows above groups | Excel product subtotal Aspose | cells.Subtotal example | group by column I | subtotal on column J | C# Excel automation
// Common Searches: Aspose.Cells Subtotal with Product function C# | place subtotal rows on top of groups Aspose.Cells | how to calculate product subtotal in Excel using Aspose | C# code for grouping and product subtotal in Aspose.Cells | cells.Subtotal summary rows above data
// Developer Intent: Insert product subtotals for column J and position the summary rows above each grouped section.
// Use Cases: Sales report that multiplies quantities per category and shows the product before the category rows. | Inventory sheet grouping items by warehouse and displaying the product of stock levels at the top of each group. | Financial ledger that groups transactions by type and adds a product subtotal row above each type for quick reference.
// AI Prompts: Show how to change the SubtotalProductDemo to use the Sum function and place subtotals below the groups. | Provide C# code that adds both Sum and Average subtotals with summary rows at the top using Aspose.Cells. | Explain how to apply Cells.Subtotal to a dynamic range when the row count is unknown.

using System;
using Aspose.Cells;

// Creates a workbook, fills column I with group keys and column J with numeric values, defines a CellArea, and calls Cells.Subtotal to group by column I, calculate the Product of column J, and insert the subtotal rows above each group before saving as SubtotalProductDemo.xlsx.
class SubtotalProductDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: column I (index 8) as group key, column J (index 9) as values
            // Header row
            cells[0, 8].PutValue("Group");
            cells[0, 9].PutValue("Amount");

            // Sample data rows
            object[,] data = new object[,]
            {
                { "A", 2 },
                { "A", 3 },
                { "B", 4 },
                { "B", 5 },
                { "C", 6 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
            {
                cells[r + 1, 8].PutValue(data[r, 0]); // Group column (I)
                cells[r + 1, 9].PutValue(data[r, 1]); // Amount column (J)
            }

            // Define the range that contains the data (including header)
            // StartRow = 0, StartColumn = 8 (I), EndRow = last data row index, EndColumn = 9 (J)
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 8,
                EndRow = data.GetLength(0), // last row index (header + data rows)
                EndColumn = 9
            };

            // Apply subtotal:
            // - Group by the first column of the range (index 0 within the range)
            // - Use Product function on the second column of the range (index 1)
            // - Place summary rows above the group data (summaryBelowData = false)
            cells.Subtotal(
                area,
                0,                                 // groupBy column index within the range
                ConsolidationFunction.Product,     // Product function
                new int[] { 1 },                   // subtotal on second column of the range
                false,                             // replace existing subtotals
                false,                             // no page breaks between groups
                false                              // summary rows at top (above data)
            );

            // Save the workbook
            workbook.Save("SubtotalProductDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
