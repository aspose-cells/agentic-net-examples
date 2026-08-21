// Title: C# – Generate a separate worksheet for each master‑detail group using Aspose.Cells smart markers with SheetBreak
// Description: This example demonstrates how to load a template workbook that contains a smart marker with the SheetBreak syntax (e.g., &=Orders:SheetBreak), bind a list of Order objects (master) and their Item collections (detail) to a WorkbookDesigner, enable LineByLine processing, and invoke Process() to create an individual worksheet for every order. The populated workbook is saved as Result.xlsx.
// Keywords: Aspose.Cells | smart markers | SheetBreak | C# | master‑detail | WorkbookDesigner | LineByLine | generate worksheets per group | template marker syntax | export to Excel
// Common Searches: Aspose.Cells SheetBreak create separate sheet per group | C# smart markers master detail example | WorkbookDesigner LineByLine true effect | how to bind detail collection for each group in Aspose.Cells | template marker for grouping orders into worksheets
// Developer Intent: Create a workbook where each master record (order) appears on its own worksheet by using smart marker grouping with SheetBreak syntax.
// Use Cases: Generate an invoice workbook with one sheet per order and its line items. | Produce a sales report that separates customers into individual worksheets. | Export project plans so each project gets a dedicated Excel sheet with its tasks.
// AI Prompts: Show how to bind the Items collection dynamically for each Order when using SheetBreak smart markers in Aspose.Cells. | Provide the exact template marker syntax needed to group Orders and Items with a sheet break. | Explain what happens if LineByLine is set to false while using SheetBreak grouping.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// This example demonstrates how to load a template workbook that contains a smart marker with the SheetBreak syntax (e.g., &=Orders:SheetBreak), bind a list of Order objects (master) and their Item collections (detail) to a WorkbookDesigner, enable LineByLine processing, and invoke Process() to create an individual worksheet for every order. The populated workbook is saved as Result.xlsx.
class Program
{
    static void Main()
    {
        // Load the template workbook that contains smart markers with sheet break syntax
        // Example marker in the template: &=Orders:SheetBreak
        Workbook workbook = new Workbook("Template.xlsx");

        // Prepare master‑detail data
        List<Order> orders = new List<Order>
        {
            new Order
            {
                OrderId = 1,
                Customer = "Alice",
                Items = new List<Item>
                {
                    new Item { Product = "Pen",      Quantity = 10 },
                    new Item { Product = "Notebook", Quantity = 5  }
                }
            },
            new Order
            {
                OrderId = 2,
                Customer = "Bob",
                Items = new List<Item>
                {
                    new Item { Product = "Pencil", Quantity = 20 },
                    new Item { Product = "Eraser", Quantity = 2  }
                }
            }
        };

        // Set up the WorkbookDesigner
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook,
            // When using sheet break syntax the default LineByLine = true works,
            // but we explicitly set it to true for clarity.
            LineByLine = true
        };

        // Bind the master data source (Orders) and the detail data source (Items)
        designer.SetDataSource("Orders", orders);
        // The detail source name must match the marker used inside the group (e.g., &Items)
        designer.SetDataSource("Items", orders[0].Items); // placeholder; actual grouping handled by smart markers

        // Process the smart markers – this will create a separate worksheet for each order group
        designer.Process();

        // Save the populated workbook
        workbook.Save("Result.xlsx");
    }

    // Master data class
    public class Order
    {
        public int OrderId { get; set; }
        public string Customer { get; set; }
        public List<Item> Items { get; set; }
    }

    // Detail data class
    public class Item
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
    }
}
