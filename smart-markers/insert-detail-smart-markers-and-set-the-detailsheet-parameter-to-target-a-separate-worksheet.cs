// Title: C# – Insert Detail Smart Markers on a Separate Worksheet Using Aspose.Cells
// Description: Creates a workbook with a master sheet and a detail sheet, defines the detail block with the _CellsSmartMarkers named range, binds nested Order data, processes master‑detail smart markers via WorkbookDesigner, and saves the result as an Excel file.
// Keywords: Aspose.Cells | C# smart markers | detail smart markers | master‑detail Excel | WorkbookDesigner | named range _CellsSmartMarkers | nested collections export | Excel report generation
// Common Searches: Aspose.Cells detail smart markers separate worksheet | C# master detail smart markers Aspose.Cells example | how to use _CellsSmartMarkers named range | populate Excel with nested objects using Aspose.Cells | WorkbookDesigner master detail report C#
// Developer Intent: Generate an Excel workbook where master smart markers reside on one sheet and detail smart markers are processed on another sheet using a named range.
// Use Cases: Invoice generation with order header on a master sheet and line items on a detail sheet. | Sales reporting that shows a summary per order and expands each order into product rows on a separate worksheet. | Exporting hierarchical data such as categories and products, placing categories on a master sheet and product listings on a linked detail sheet.
// AI Prompts: Show how to set WorkbookDesigner.DetailSheetName explicitly instead of relying on the _CellsSmartMarkers range. | Provide code to apply formatting (bold headers, borders) to the detail rows after smart marker processing. | Explain how to bind multiple data sources (e.g., Customers and Orders) while using master‑detail smart markers across different worksheets.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a workbook with a master sheet and a detail sheet, defines the detail block with the _CellsSmartMarkers named range, binds nested Order data, processes master‑detail smart markers via WorkbookDesigner, and saves the result as an Excel file.
public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            DetailSmartMarkersExample.Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class DetailSmartMarkersExample
{
    public static void Run()
    {
        try
        {
            // 1. Create a new workbook (template)
            Workbook wb = new Workbook();

            // -----------------------------------------------------------------
            // 2. Prepare worksheets
            // -----------------------------------------------------------------
            // Master worksheet (index 0) – will contain master smart markers
            Worksheet master = wb.Worksheets[0];
            master.Name = "Master";

            // Detail worksheet – will contain detail smart markers
            Worksheet detail = wb.Worksheets.Add("Detail");

            // -----------------------------------------------------------------
            // 3. Insert master smart markers
            // -----------------------------------------------------------------
            master.Cells["A1"].PutValue("Order ID");
            master.Cells["B1"].PutValue("Customer");
            master.Cells["A2"].PutValue("&=Orders.OrderID");
            master.Cells["B2"].PutValue("&=Orders.CustomerName");

            // -----------------------------------------------------------------
            // 4. Insert detail smart markers
            // -----------------------------------------------------------------
            // Define a named range that Aspose.Cells recognises as the detail block
            AsposeRange detailRange = detail.Cells.CreateRange("A1:B1");
            detailRange.Name = "_CellsSmartMarkers";

            // Header row in the detail sheet
            detail.Cells["A1"].PutValue("Product");
            detail.Cells["B1"].PutValue("Quantity");

            // Data row – detail smart markers that repeat for each order detail
            detail.Cells["A2"].PutValue("&=Orders.OrderDetails.ProductName");
            detail.Cells["B2"].PutValue("&=Orders.OrderDetails.Quantity");

            // -----------------------------------------------------------------
            // 5. Prepare sample data
            // -----------------------------------------------------------------
            List<Order> orders = new List<Order>
            {
                new Order
                {
                    OrderID = 1001,
                    CustomerName = "John Doe",
                    OrderDetails = new List<OrderDetail>
                    {
                        new OrderDetail { ProductName = "Pen", Quantity = 10 },
                        new OrderDetail { ProductName = "Notebook", Quantity = 5 }
                    }
                },
                new Order
                {
                    OrderID = 1002,
                    CustomerName = "Jane Smith",
                    OrderDetails = new List<OrderDetail>
                    {
                        new OrderDetail { ProductName = "Pencil", Quantity = 20 }
                    }
                }
            };

            // -----------------------------------------------------------------
            // 6. Configure WorkbookDesigner
            // -----------------------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb
                // No need to set DetailSheetName; the named range "_CellsSmartMarkers" identifies the detail sheet.
            };

            // Bind the data source (the name "Orders" matches the smart markers)
            designer.SetDataSource("Orders", orders);

            // Process all smart markers (master and detail)
            designer.Process();

            // -----------------------------------------------------------------
            // 7. Save the populated workbook
            // -----------------------------------------------------------------
            string outputPath = "DetailSmartMarkersOutput.xlsx";

            // Ensure the directory exists before saving
            string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                throw;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Run failed: {ex.Message}");
            throw;
        }
    }

    // -----------------------------------------------------------------
    // Data model classes used as the data source
    // -----------------------------------------------------------------
    public class Order
    {
        public int OrderID { get; set; }
        public string CustomerName { get; set; } = string.Empty;
        public List<OrderDetail> OrderDetails { get; set; } = new List<OrderDetail>();
    }

    public class OrderDetail
    {
        public string ProductName { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }
}
