// Title: Import a specific list element by index using Aspose.Cells smart markers in C#
// Common Searches: Aspose.Cells smart marker syntax to access the third item in a C# list | How to bind a List<T> to WorkbookDesigner and reference an element by index | C# example of using &=Orders[2].ItemName in an Excel template with Aspose.Cells | Retrieve a single collection element in Excel export using Aspose.Cells smart markers | Index-based smart marker for array element in Aspose.Cells C# tutorial
// Tags: smart markers indexed collection binding | Aspose.Cells WorkbookDesigner list data source | C# export specific list element to Excel | smart marker array index syntax | Aspose.Cells retrieve collection element by position

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // Simple POCO representing an order
    // The example creates a List<Order>, inserts the smart marker "&=Orders[2].ItemName" into cell A1, binds the list to a WorkbookDesigner, processes the smart markers, and saves the workbook as Output.xlsx, demonstrating how to import a specific list element by its index using Aspose.Cells smart markers in C#.
    public class Order
    {
        public string ItemName { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare sample data
                List<Order> orders = new List<Order>
                {
                    new Order { ItemName = "Apple",  Quantity = 10, Price = 0.5 },
                    new Order { ItemName = "Banana", Quantity = 20, Price = 0.3 },
                    new Order { ItemName = "Cherry", Quantity = 15, Price = 0.8 }, // index 2
                    new Order { ItemName = "Date",   Quantity = 5,  Price = 1.2 }
                };

                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Insert a smart marker that references the third element (index 2) of the Orders list
                // The syntax "&=Orders[2].ItemName" tells Aspose.Cells to pull ItemName from orders[2]
                sheet.Cells["A1"].PutValue("&=Orders[2].ItemName");

                // Set up the designer and bind the data source
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Orders", orders);
                designer.Process(); // Process smart markers

                // Define output file path
                string outputPath = "Output.xlsx";

                // Save the result
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
