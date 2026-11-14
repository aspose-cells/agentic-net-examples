// Title: Import a specific list element by index using Aspose.Cells smart markers in C#
// Common Searches: Aspose.Cells smart marker how to reference the third item in a list | C# syntax for importing a specific collection element using smart markers | Using WorkbookDesigner to set a List<T> data source for indexed smart markers | Example of smart marker '&=Orders[2].ItemName' to populate an Excel cell | How to process smart markers that include an array index in Aspose.Cells
// Tags: smart markers indexed collection import | Aspose.Cells WorkbookDesigner list data source | C# import specific list element into Excel | smart marker array element syntax | process smart markers with indexed reference

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates placing the smart marker "&=Orders[2].ItemName" in a worksheet cell, setting a List<Order> as the data source named "Orders" with WorkbookDesigner, processing the marker to import the ItemName of the third list element, and saving the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a smart marker that references the third element (index 2) of the Orders collection
        // The syntax &=Orders[2].ItemName tells Aspose.Cells to import the ItemName of that element
        sheet.Cells["A1"].PutValue("&=Orders[2].ItemName");

        // Prepare the data source: a list of Order objects
        List<Order> orders = new List<Order>
        {
            new Order { ItemName = "Apple" },
            new Order { ItemName = "Banana" },
            new Order { ItemName = "Cherry" }, // This is the element at index 2
            new Order { ItemName = "Date" }
        };

        // Set the data source for the smart marker processor
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Orders", orders);

        // Process the smart markers and populate the worksheet
        designer.Process();

        // Save the resulting workbook
        workbook.Save("Output.xlsx");
    }
}

// Simple POCO representing an order
public class Order
{
    public string ItemName { get; set; }
}
