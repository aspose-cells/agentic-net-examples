// Title: Use a custom DataTable with WorkbookDesigner to populate smart markers in Aspose.Cells for .NET (C#)
// Description: This C# example shows how to create a workbook, define smart markers that reference a runtime DataTable named "Products", fill the table with data, assign it to WorkbookDesigner via SetDataSource, process the markers, and save the populated Excel file.
// Keywords: Aspose.Cells | WorkbookDesigner | SetDataSource | DataTable | smart markers | C# .NET | Excel generation | custom data source | template population | example code
// Common Searches: Aspose.Cells WorkbookDesigner SetDataSource C# | bind DataTable to smart markers Aspose | populate Excel template from DataTable .NET | smart markers custom data source example | how to use WorkbookDesigner with DataTable
// Developer Intent: Bind a DataTable to WorkbookDesigner, process smart markers, and generate a filled Excel workbook.
// Use Cases: Generate a product catalog by filling smart markers from a DataTable. | Create invoices where line‑item details are supplied by a DataTable. | Build multi‑sheet reports, each sheet driven by a different DataTable. | Export query results from a database to an Excel template. | Automate data‑driven dashboards using smart markers and runtime data.
// AI Prompts: Provide C# code that loads an existing .xlsx template, sets a DataSet with several tables as the data source for WorkbookDesigner, and processes smart markers for each sheet. | Explain step‑by‑step how WorkbookDesigner.SetDataSource works with a DataTable and how to reference its columns in smart markers. | Generate a sample that uses a DataTable of employee records to fill a payroll worksheet via smart markers.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsCustomDataSourceDemo
{
    // This C# example shows how to create a workbook, define smart markers that reference a runtime DataTable named "Products", fill the table with data, assign it to WorkbookDesigner via SetDataSource, process the markers, and save the populated Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (template can be loaded here if needed)
            Workbook workbook = new Workbook();

            // Add a worksheet and place smart markers where data will be inserted
            Worksheet sheet = workbook.Worksheets[0];
            // Smart markers use the name of the DataTable ("Products") followed by column names
            sheet.Cells["A1"].PutValue("Product ID");
            sheet.Cells["B1"].PutValue("Product Name");
            sheet.Cells["C1"].PutValue("Price");
            // Markers start from row 2
            sheet.Cells["A2"].PutValue("&=Products.ProductID");
            sheet.Cells["B2"].PutValue("&=Products.ProductName");
            sheet.Cells["C2"].PutValue("&=Products.Price");

            // Create a custom DataTable and populate it with sample data
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("ProductID", typeof(int));
            dataTable.Columns.Add("ProductName", typeof(string));
            dataTable.Columns.Add("Price", typeof(decimal));

            dataTable.Rows.Add(101, "Laptop", 1200.50m);
            dataTable.Rows.Add(102, "Smartphone", 799.99m);
            dataTable.Rows.Add(103, "Tablet", 450.00m);

            // Initialize WorkbookDesigner and assign the workbook
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // Set the custom DataTable as the data source (uses SetDataSource(DataTable) rule)
            designer.SetDataSource(dataTable);

            // Process the smart markers to populate the worksheet with data
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("CustomDataSourceOutput.xlsx");
        }
    }
}
