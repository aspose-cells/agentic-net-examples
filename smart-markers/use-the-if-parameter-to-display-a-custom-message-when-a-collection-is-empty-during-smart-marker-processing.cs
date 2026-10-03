// Title: Display a custom message for an empty DataTable using the If parameter in Aspose.Cells smart markers (C#)
// AI Prompts: Generate C# code that inserts a smart marker with the &=If syntax to show "No orders found." when the bound Orders table has zero rows, then processes it with WorkbookDesigner. | Create an Excel template, place a conditional smart marker that checks Orders.Count=0, bind an empty DataSet, and save the output workbook using Aspose.Cells for .NET. | Write a C# example demonstrating how to use the If parameter in smart markers to handle empty collections and output a fallback message.
// Common Searches: Aspose.Cells C# smart marker show message when DataTable is empty | How to use &=If in Aspose.Cells smart markers for zero‑record collections | WorkbookDesigner conditional smart marker syntax for empty dataset | Display "No data" text in Excel using Aspose.Cells smart markers and If parameter | C# example of smart marker If parameter with empty Orders table
// Tags: Aspose.Cells smart marker If parameter | C# WorkbookDesigner empty collection handling | Excel template conditional text with smart markers | Aspose.Cells display no‑data message | smart marker collection count check C#

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads a template workbook, inserts a smart marker using the If parameter to output "No orders found." when the Orders DataTable is empty, binds an empty DataSet as the data source, processes the markers with WorkbookDesigner, and saves the resulting Excel file.
class SmartMarkerIfExample
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);

            // Access the first worksheet where the smart marker is placed
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker that uses the If parameter.
            // Syntax: &=If(Orders.Count=0, "No orders found.", "")
            sheet.Cells["A1"].PutValue("&=If(Orders.Count=0, \"No orders found.\", \"\")");

            // Prepare the data source: an empty DataTable named "Orders"
            DataTable ordersTable = new DataTable("Orders");
            ordersTable.Columns.Add("OrderID", typeof(int));
            ordersTable.Columns.Add("Product", typeof(string));
            ordersTable.Columns.Add("Quantity", typeof(int));
            // No rows are added, so the collection is empty

            // Create a DataSet and add the empty table
            DataSet data = new DataSet();
            data.Tables.Add(ordersTable);

            // Process the smart markers using WorkbookDesigner (correct API)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(data);
            designer.Process();

            // Save the result workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
