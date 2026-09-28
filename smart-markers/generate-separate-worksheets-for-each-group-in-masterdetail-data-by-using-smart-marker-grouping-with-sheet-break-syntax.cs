// Title: Create separate worksheets for each order using Aspose.Cells sheet‑break smart marker in C#
// AI Prompts: Build an Excel template that places the sheet‑break smart marker (&=Orders:SheetBreak) followed by detail markers, bind a List<Order> to the designer, run Process, and generate one worksheet per order. | Write C# code to define master‑detail classes, insert sheet‑break and item markers, invoke the Aspose.Cells designer to populate the workbook, and save the result as an .xlsx file.
// Common Searches: Aspose.Cells C# sheet break smart marker create worksheet per collection item | How to generate separate Excel sheets for each order using master‑detail smart markers | Split an Excel workbook into multiple sheets with sheet break syntax in Aspose.Cells | C# example of grouping data with smart markers and sheet break for master‑detail
// Tags: Aspose.Cells sheet break marker | C# generate separate worksheets from collection | Excel .xlsx sheets per order using smart markers | smart marker grouping with sheet break | export order data to individual worksheets

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerSheetBreakDemo
{
    // Master class representing a group (e.g., an order)
    // The sample creates a workbook template, adds a sheet‑break smart marker (&=Orders:SheetBreak) and detail markers, binds a List<Order> (master‑detail data) to the Aspose.Cells designer, processes the markers to produce a distinct worksheet for each order, and saves the file as MasterDetail_SheetBreak_Output.xlsx.
    public class Order
    {
        public int OrderId { get; set; }
        public string Customer { get; set; }
        public List<OrderItem> Items { get; set; }
    }

    // Detail class representing items within a group
    public class OrderItem
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -------------------------------------------------
            // 1. Prepare master‑detail data
            // -------------------------------------------------
            var orders = new List<Order>
            {
                new Order
                {
                    OrderId = 1001,
                    Customer = "Alice",
                    Items = new List<OrderItem>
                    {
                        new OrderItem { Product = "Apple",  Quantity = 5 },
                        new OrderItem { Product = "Banana", Quantity = 3 }
                    }
                },
                new Order
                {
                    OrderId = 1002,
                    Customer = "Bob",
                    Items = new List<OrderItem>
                    {
                        new OrderItem { Product = "Orange", Quantity = 2 },
                        new OrderItem { Product = "Grapes", Quantity = 4 },
                        new OrderItem { Product = "Mango",  Quantity = 1 }
                    }
                }
            };

            // -------------------------------------------------
            // 2. Create a workbook template with smart markers
            // -------------------------------------------------
            Workbook wb = new Workbook();                     // create workbook
            Worksheet ws = wb.Worksheets[0];                  // access first sheet
            Cells cells = ws.Cells;

            // Header for master data
            cells["A1"].PutValue("Order ID");
            cells["B1"].PutValue("Customer");

            // Sheet break smart marker – creates a new sheet for each Order
            // The marker must be placed in a cell that will be processed first.
            cells["A2"].PutValue("&=Orders:SheetBreak");

            // Header for detail data (will appear on each generated sheet)
            cells["A3"].PutValue("Product");
            cells["B3"].PutValue("Quantity");

            // Detail smart markers – repeat for each item in the current Order
            cells["A4"].PutValue("&=Orders.Items.Product");
            cells["B4"].PutValue("&=Orders.Items.Quantity");

            // -------------------------------------------------
            // 3. Set up WorkbookDesigner and bind data source
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb
            };
            // Bind the master collection to the name used in smart markers
            designer.SetDataSource("Orders", orders);

            // -------------------------------------------------
            // 4. Process the smart markers (creates separate sheets)
            // -------------------------------------------------
            designer.Process();   // processes all smart markers in the workbook

            // -------------------------------------------------
            // 5. Save the result
            // -------------------------------------------------
            wb.Save("MasterDetail_SheetBreak_Output.xlsx");
        }
    }
}
