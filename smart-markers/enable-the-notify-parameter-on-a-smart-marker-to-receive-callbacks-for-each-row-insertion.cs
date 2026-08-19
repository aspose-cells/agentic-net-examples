// Title: Aspose.Cells for .NET – Using ?Notify on Smart Markers with ISmartMarkerCallBack to Capture Row‑Insertion Events (C#)
// Description: Demonstrates how to add the ?Notify suffix to smart markers, implement ISmartMarkerCallBack, and configure WorkbookDesigner so that a callback is fired for every row generated during smart‑marker processing. The example creates a workbook, defines a DataTable, registers the callback, processes the markers, and saves the result.
// Keywords: Aspose.Cells smart marker notify | C# ?Notify callback | ISmartMarkerCallBack example | WorkbookDesigner row insertion event | smart marker per‑row notification | Aspose.Cells .NET sample | smart marker processing callback | Aspose.Cells GitHub example
// Common Searches: Aspose.Cells ?Notify smart marker C# | How to receive row callbacks with smart markers | Implement ISmartMarkerCallBack in Aspose.Cells | Enable per‑row notifications for smart markers | Aspose.Cells smart marker callback example
// Developer Intent: Add the ?Notify parameter to smart markers and handle each generated row via ISmartMarkerCallBack in C#.
// Use Cases: Log or audit every row inserted during smart‑marker expansion. | Trigger UI updates or external service calls for each new row. | Validate or transform data on‑the‑fly while populating a workbook.
// AI Prompts: Generate a version of MySmartMarkerCallback that stores row indices in a List<int> instead of printing to the console. | Show how to use ?Notify with multiple smart‑marker tables and differentiate callbacks by table name. | Explain how to stop the callback after processing a specific number of rows in the ISmartMarkerCallBack implementation.

using System;
using System.Data;
using Aspose.Cells;

// Implement the callback interface to receive notifications for each row insertion
// Demonstrates how to add the ?Notify suffix to smart markers, implement ISmartMarkerCallBack, and configure WorkbookDesigner so that a callback is fired for every row generated during smart‑marker processing. The example creates a workbook, defines a DataTable, registers the callback, processes the markers, and saves the result.
public class MySmartMarkerCallback : ISmartMarkerCallBack
{
    // This method is called by Aspose.Cells for each smart marker row that is processed
    public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
    {
        Console.WriteLine($"Callback - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");
    }
}

public class EnableSmartMarkerNotifyDemo
{
    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a smart marker with the Notify parameter.
        // The "?Notify" suffix tells Aspose.Cells to invoke the callback for each row insertion.
        // Example smart marker: &=$Products.ProductName?Notify
        sheet.Cells["A1"].PutValue("&=$Products.ProductName?Notify");
        sheet.Cells["B1"].PutValue("&=$Products.Price?Notify");

        // Prepare a DataTable as the data source with multiple rows
        DataTable dt = new DataTable("Products");
        dt.Columns.Add("ProductName", typeof(string));
        dt.Columns.Add("Price", typeof(double));

        dt.Rows.Add("Apple", 1.20);
        dt.Rows.Add("Banana", 0.80);
        dt.Rows.Add("Cherry", 2.50);

        // Set up the WorkbookDesigner
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook,
            // Assign the callback implementation
            CallBack = new MySmartMarkerCallback()
        };

        // Register the data source
        designer.SetDataSource(dt);

        // Process the smart markers. The callback will be triggered for each inserted row.
        designer.Process(false);

        // Save the resulting workbook
        workbook.Save("SmartMarkerNotifyResult.xlsx");
    }
}

// Entry point
class Program
{
    static void Main()
    {
        EnableSmartMarkerNotifyDemo.Run();
    }
}
