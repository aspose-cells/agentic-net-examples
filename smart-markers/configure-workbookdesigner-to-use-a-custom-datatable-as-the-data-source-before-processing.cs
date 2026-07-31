// Title: C# – Use WorkbookDesigner with a custom DataTable to fill smart markers in Aspose.Cells
// Description: Shows how to create a Workbook, insert smart markers that reference a DataTable named Products, build the table in code, bind it to WorkbookDesigner via SetDataSource, process the markers, and save the populated Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | WorkbookDesigner | SetDataSource | DataTable | smart markers | C# .NET | Excel export | template filling | custom data source | product catalog example
// Common Searches: WorkbookDesigner SetDataSource DataTable example | Aspose.Cells populate smart markers from DataTable | C# fill Excel template with DataTable | How to bind DataTable to WorkbookDesigner | Aspose.Cells smart markers tutorial
// Developer Intent: Bind a user‑created DataTable to WorkbookDesigner so that smart markers in an Excel template are automatically populated.
// Use Cases: Generate a product catalog Excel file by linking a Products DataTable to smart markers. | Create invoices by mapping an Orders DataTable to markers in a pre‑designed worksheet. | Export a financial summary by associating a ReportData DataTable with smart markers in a template.
// AI Prompts: Provide C# code to attach multiple DataTables to a single WorkbookDesigner, each using a distinct smart‑marker prefix. | Show how to load an existing Excel template and set a DataSet as the data source for WorkbookDesigner smart markers. | Explain strategies for handling null or missing values in a DataTable when processing smart markers with WorkbookDesigner.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a Workbook, insert smart markers that reference a DataTable named Products, build the table in code, bind it to WorkbookDesigner via SetDataSource, process the markers, and save the populated Excel file using Aspose.Cells for .NET.
    class WorkbookDesignerCustomDataTableDemo
    {
        static void Main()
        {
            // Create a new workbook (template can be loaded here if needed)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add smart markers where data will be inserted
            Worksheet sheet = workbook.Worksheets[0];
            // Smart markers use the name of the DataTable (e.g., "Products")
            sheet.Cells["A1"].PutValue("Product ID");
            sheet.Cells["B1"].PutValue("Product Name");
            sheet.Cells["C1"].PutValue("Price");
            // Markers for data rows
            sheet.Cells["A2"].PutValue("&=$Products.ProductID");
            sheet.Cells["B2"].PutValue("&=$Products.ProductName");
            sheet.Cells["C2"].PutValue("&=$Products.Price");

            // Build a custom DataTable with sample data
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("ProductID", typeof(int));
            dataTable.Columns.Add("ProductName", typeof(string));
            dataTable.Columns.Add("Price", typeof(decimal));

            dataTable.Rows.Add(101, "Laptop", 1200.50m);
            dataTable.Rows.Add(102, "Smartphone", 799.99m);
            dataTable.Rows.Add(103, "Tablet", 450.00m);

            // Initialize the WorkbookDesigner and assign the workbook
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // Set the custom DataTable as the data source (uses SetDataSource(DataTable) rule)
            designer.SetDataSource(dataTable);

            // Process the smart markers to populate the worksheet with the data
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("CustomDataTableOutput.xlsx");
        }
    }
}
