// Title: C# – Group Master‑Detail Rows with Aspose.Cells Smart Markers (Parent Above Child)
// Description: Creates an Excel workbook, sets a parent smart marker above a child range, processes a hierarchical List<Order> with WorkbookDesigner, then uses Cells.GroupRows to collapse each order's detail rows for a tidy master‑detail view.
// Keywords: Aspose.Cells | C# | smart markers | master detail | row grouping | hierarchical data source | WorkbookDesigner | Excel export | parent smart marker | child range
// Common Searches: Aspose.Cells group master detail rows C# | smart markers parent above child range | process child range only Aspose.Cells | how to collapse detail rows with Aspose.Cells | hierarchical data source smart markers example
// Developer Intent: Generate an Excel file that lists orders with expandable/collapsible detail rows by using smart markers and row grouping in C#.
// Use Cases: Invoice reports where each invoice header can expand to show line items. | Sales shipment sheets that group product shipments under each shipment ID. | Project task lists with subtasks grouped under their parent tasks for easy navigation.
// AI Prompts: Write C# code using Aspose.Cells to create a master‑detail Excel sheet with a parent smart marker above the child range and group the detail rows. | Explain how to bind a hierarchical List<Order> to smart markers, process only the child range, and then apply row grouping in Aspose.Cells. | Provide step‑by‑step instructions to add a subtotal row for each order while keeping the detail rows grouped with smart markers.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsMasterDetailExample
{
    // Master class
    // Creates an Excel workbook, sets a parent smart marker above a child range, processes a hierarchical List<Order> with WorkbookDesigner, then uses Cells.GroupRows to collapse each order's detail rows for a tidy master‑detail view.
    public class Order
    {
        public int OrderID { get; set; }
        public DateTime OrderDate { get; set; }
        public List<Detail> Details { get; set; } = new List<Detail>();
    }

    // Detail class
    public class Detail
    {
        public string Product { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ----- Set up smart markers -----
                // Parent (master) smart markers
                sheet.Cells["A1"].PutValue("&=Orders.OrderID");
                sheet.Cells["B1"].PutValue("&=Orders.OrderDate");

                // Child (detail) smart markers – placed below the parent row
                sheet.Cells["A2"].PutValue("&=Orders.Details.Product");
                sheet.Cells["B2"].PutValue("&=Orders.Details.Quantity");

                // Define a range that contains the child markers.
                // This range will be processed repeatedly for each master record.
                Aspose.Cells.Range childRange = sheet.Cells.CreateRange("A2:B5");
                childRange.Name = "_CellsSmartMarkers";

                // ----- Prepare hierarchical data source -----
                List<Order> orders = new List<Order>
                {
                    new Order
                    {
                        OrderID = 1001,
                        OrderDate = new DateTime(2023, 1, 10),
                        Details = new List<Detail>
                        {
                            new Detail { Product = "Apple",  Quantity = 10 },
                            new Detail { Product = "Banana", Quantity = 5 }
                        }
                    },
                    new Order
                    {
                        OrderID = 1002,
                        OrderDate = new DateTime(2023, 2, 15),
                        Details = new List<Detail>
                        {
                            new Detail { Product = "Orange", Quantity = 8 },
                            new Detail { Product = "Grapes", Quantity = 12 },
                            new Detail { Product = "Mango",  Quantity = 7 }
                        }
                    }
                };

                // ----- Apply smart markers -----
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                // Set the hierarchical data source
                designer.SetDataSource("Orders", orders);
                // Process only the child range (true = process this range only)
                designer.Process(childRange, true);

                // ----- Group detail rows under each master row -----
                // After processing, the rows are laid out as:
                // Row 0 : Master 1
                // Row 1-2 : Details of Master 1
                // Row 3 : Master 2
                // Row 4-6 : Details of Master 2
                // Group rows 1-2 (detail of first order)
                sheet.Cells.GroupRows(1, 2, true);
                // Group rows 4-6 (detail of second order)
                sheet.Cells.GroupRows(4, 6, true);

                // Save the result
                string outputPath = "MasterDetailGrouped.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
