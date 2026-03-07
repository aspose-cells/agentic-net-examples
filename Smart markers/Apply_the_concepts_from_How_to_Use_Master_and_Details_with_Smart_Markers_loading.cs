using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Determine the template file path
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string templatePath = Path.Combine(baseDir, "MasterDetailTemplate.xlsx");

        // If the template does not exist, create a minimal one
        if (!File.Exists(templatePath))
        {
            var tempWb = new Workbook();
            var tempSheet = tempWb.Worksheets[0];

            // Simple smart markers for demonstration
            tempSheet.Cells["A2"].PutValue("&=Orders.OrderID");
            tempSheet.Cells["B2"].PutValue("&=Orders.CustomerName");
            tempSheet.Cells["C2"].PutValue("&=Orders.OrderDate");

            tempSheet.Cells["E2"].PutValue("&=Orders.Details.Product");
            tempSheet.Cells["F2"].PutValue("&=Orders.Details.Quantity");
            tempSheet.Cells["G2"].PutValue("&=Orders.Details.Price");

            // Define master and detail ranges for range smart markers
            AsposeRange masterRange = tempSheet.Cells.CreateRange("A2:C2");
            masterRange.Name = "_CellsSmartMarkers";

            AsposeRange detailRange = tempSheet.Cells.CreateRange("E2:G2");
            detailRange.Name = "_CellsSmartMarkers";

            tempWb.Save(templatePath);
        }

        // Load the template workbook that contains smart markers
        Workbook workbook = new Workbook(templatePath);
        Worksheet sheet = workbook.Worksheets[0];

        // Sample master‑detail data
        List<Order> orders = new List<Order>
        {
            new Order
            {
                OrderID = 1001,
                CustomerName = "John Doe",
                OrderDate = new DateTime(2023, 1, 10),
                Details = new List<OrderDetail>
                {
                    new OrderDetail { Product = "Laptop", Quantity = 1, Price = 1200.00 },
                    new OrderDetail { Product = "Mouse", Quantity = 2, Price = 25.00 }
                }
            },
            new Order
            {
                OrderID = 1002,
                CustomerName = "Alice Smith",
                OrderDate = new DateTime(2023, 2, 5),
                Details = new List<OrderDetail>
                {
                    new OrderDetail { Product = "Phone", Quantity = 1, Price = 800.00 },
                    new OrderDetail { Product = "Headset", Quantity = 1, Price = 60.00 }
                }
            }
        };

        // Initialize WorkbookDesigner and assign the workbook
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };

        // Set the master data source; detail data is accessed via the property name "Details"
        designer.SetDataSource("Orders", orders);

        // Process all smart markers (master and detail)
        designer.Process();

        // Save the populated workbook
        string outputPath = Path.Combine(baseDir, "MasterDetailOutput.xlsx");
        workbook.Save(outputPath);
    }

    // Master data class
    public class Order
    {
        public int OrderID { get; set; }
        public string CustomerName { get; set; } = null!;
        public DateTime OrderDate { get; set; }
        public List<OrderDetail> Details { get; set; } = null!;
    }

    // Detail data class
    public class OrderDetail
    {
        public string Product { get; set; } = null!;
        public int Quantity { get; set; }
        public double Price { get; set; }
    }
}