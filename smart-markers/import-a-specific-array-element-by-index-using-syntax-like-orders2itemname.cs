// Title: Import a specific list element by index using Aspose.Cells smart markers in C# (e.g., &=Orders[2].ItemName)
// AI Prompts: Write C# code that inserts the smart marker '&=Orders[2].ItemName' into cell A1, binds a List<Order> to a DataSet, and processes the marker with WorkbookDesigner. | Show how to create a DataSet from a collection and use a smart marker to populate an Excel cell with the third item's ItemName. | Demonstrate saving the workbook after processing indexed smart markers to a file named Result.xlsx.
// Common Searches: Aspose.Cells how to reference the third element of a list in a smart marker | C# smart marker syntax for accessing Orders[2] in Excel | Binding a List<T> to WorkbookDesigner for indexed smart marker values | Using &=Orders[2].ItemName to pull a single record into an Excel cell | Example of processing smart markers with a DataSet source in Aspose.Cells
// Tags: smart marker collection index Aspose.Cells | WorkbookDesigner DataSet binding C# | reference list element in Excel smart marker | export workbook after processing smart markers | Aspose.Cells indexed marker example

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

// The example creates a workbook, places the smart marker '&=Orders[2].ItemName' in cell A1, converts a List<Order> into a DataSet, sets the DataSet as the data source for WorkbookDesigner, processes the marker to insert the third order's ItemName, and saves the result as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker that references the third element (index 2) of the Orders collection
            // The syntax &=Orders[2].ItemName tells Aspose.Cells to pull ItemName from Orders[2]
            sheet.Cells["A1"].PutValue("&=Orders[2].ItemName");

            // Prepare the data source with an Orders collection
            var orders = new List<Order>
            {
                new Order { ItemName = "Apple",  Quantity = 5 },
                new Order { ItemName = "Banana", Quantity = 10 },
                new Order { ItemName = "Cherry", Quantity = 15 }, // Index 2
                new Order { ItemName = "Date",   Quantity = 20 }
            };

            // Convert the list to a DataSet (required by WorkbookDesigner in this context)
            DataTable ordersTable = new DataTable("Orders");
            ordersTable.Columns.Add("ItemName", typeof(string));
            ordersTable.Columns.Add("Quantity", typeof(int));

            foreach (var o in orders)
            {
                ordersTable.Rows.Add(o.ItemName, o.Quantity);
            }

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(ordersTable);

            // Process the smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Save the resulting workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple POCO representing an order item
    public class Order
    {
        public string ItemName { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }
}
