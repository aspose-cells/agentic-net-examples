// Title: Generate a master‑detail Excel sheet by positioning a parent smart marker above a child smart marker range with Aspose.Cells for .NET
// AI Prompts: Generate an Excel file that leverages a parent smart marker positioned before a child smart marker range to render hierarchical order information using Aspose.Cells WorkbookDesigner. | Modify the smart‑marker range to include extra detail columns (e.g., price) and process only that range with Aspose.Cells. | Add a summary row after each master record by employing hierarchical smart markers with WorkbookDesigner in Aspose.Cells.
// Common Searches: how to use range smart markers for master detail in Aspose.Cells .NET | Aspose.Cells place parent smart marker above child markers example | process only a specific smart marker range with WorkbookDesigner | export hierarchical list to Excel using Aspose.Cells smart markers C# | group master detail rows in Excel using Aspose.Cells range smart markers
// Tags: smart marker hierarchy processing Aspose.Cells | WorkbookDesigner process defined range | export nested collections to Excel .NET | smart marker row grouping Excel | parent marker positioning Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsMasterDetailDemo
{
    // Data model for master‑detail
    // The example creates a workbook, defines a smart‑marker range where a parent marker (Orders.OrderID) sits above child markers (Orders.OrderDetails.Product and Quantity), binds a hierarchical List<Order> as the data source, processes only the specified range with WorkbookDesigner, and saves the resulting master‑detail Excel file.
    public class Order
    {
        public int OrderID { get; set; }
        public List<OrderDetail> OrderDetails { get; set; } = new List<OrderDetail>();
    }

    public class OrderDetail
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
                // Create a new workbook (lifecycle: create)
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // ----- Template with smart markers -----
                // Parent smart marker (master) placed above child range
                cells["A1"].PutValue("&=Orders.OrderID");               // Master marker
                cells["A2"].PutValue("&=Orders.OrderDetails.Product"); // Child marker - first column
                cells["B2"].PutValue("&=Orders.OrderDetails.Quantity"); // Child marker - second column

                // Define the range that contains the smart markers and give it the required name
                // This enables range smart markers processing
                AsposeRange smRange = cells.CreateRange("A1:B2");
                smRange.Name = "_CellsSmartMarkers";

                // ----- Prepare hierarchical data source -----
                var orders = new List<Order>
                {
                    new Order
                    {
                        OrderID = 1001,
                        OrderDetails = new List<OrderDetail>
                        {
                            new OrderDetail { Product = "Apple",  Quantity = 5 },
                            new OrderDetail { Product = "Banana", Quantity = 3 }
                        }
                    },
                    new Order
                    {
                        OrderID = 1002,
                        OrderDetails = new List<OrderDetail>
                        {
                            new OrderDetail { Product = "Orange", Quantity = 7 },
                            new OrderDetail { Product = "Grapes", Quantity = 2 },
                            new OrderDetail { Product = "Mango",  Quantity = 4 }
                        }
                    }
                };

                // ----- Process smart markers -----
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = wb
                };
                designer.SetDataSource("Orders", orders);
                // Process only the defined range (lifecycle: process)
                designer.Process(smRange, true);

                // Save the result (lifecycle: save)
                string outputPath = "MasterDetailOutput.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
