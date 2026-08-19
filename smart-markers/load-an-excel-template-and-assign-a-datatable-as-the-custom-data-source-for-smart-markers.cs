// Title: C# – Assign a DataTable to Smart Markers in an Excel Template with Aspose.Cells
// Description: Load an Excel workbook that contains smart markers, create a DataTable with product data, bind the table to WorkbookDesigner using SetDataSource, process the markers to populate the sheet, and save the result as a new file. Demonstrates a complete workflow for custom data sources in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | WorkbookDesigner | SetDataSource | Smart markers | DataTable binding | Excel template processing | populate Excel with DataTable | Aspose.Cells example | dynamic Excel report
// Common Searches: how to bind a DataTable to smart markers Aspose.Cells | WorkbookDesigner SetDataSource C# example | populate Excel template using smart markers and DataTable | Aspose.Cells process smart markers after setting data source | C# code for smart markers with custom DataTable
// Developer Intent: Load an Excel template, attach a DataTable as the smart‑marker data source, execute the markers, and write the filled workbook to disk.
// Use Cases: Generate a product catalog by feeding a Products DataTable into a smart‑marker template. | Create batch invoices where each invoice uses a DataTable of line items as the smart‑marker source. | Build sales dashboards that automatically fill charts and tables from a sales DataTable.
// AI Prompts: Write C# code that loads an Excel file with smart markers, sets a DataTable named 'Orders' as the data source via WorkbookDesigner, processes the markers, and saves the output. | Show how to assign multiple DataTables to different smart‑marker groups in Aspose.Cells and provide sample code for each assignment. | List troubleshooting steps when smart markers remain empty after calling SetDataSource and Process with a DataTable.

using System;
using System.Data;
using Aspose.Cells;

// Load an Excel workbook that contains smart markers, create a DataTable with product data, bind the table to WorkbookDesigner using SetDataSource, process the markers to populate the sheet, and save the result as a new file. Demonstrates a complete workflow for custom data sources in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the Excel template that contains smart markers
        Workbook workbook = new Workbook("Template.xlsx");

        // Initialize the WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Create a DataTable that will be used as the custom data source
        DataTable dataTable = new DataTable("Products");
        dataTable.Columns.Add("ProductID", typeof(int));
        dataTable.Columns.Add("ProductName", typeof(string));

        // Populate the DataTable with sample data
        dataTable.Rows.Add(1, "Laptop");
        dataTable.Rows.Add(2, "Smartphone");
        dataTable.Rows.Add(3, "Tablet");

        // Assign the DataTable to the designer as the data source for smart markers
        designer.SetDataSource(dataTable);

        // Process the smart markers and fill the worksheet with data
        designer.Process();

        // Save the processed workbook to a new file
        workbook.Save("Result.xlsx");
    }
}
