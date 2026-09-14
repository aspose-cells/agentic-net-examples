// Title: How to load an Excel template and bind a DataTable as a custom data source for Aspose.Cells smart markers in C#
// AI Prompts: Write C# code that opens an existing .xlsx template, creates a DataTable with product data, assigns it to WorkbookDesigner via SetDataSource, processes the smart markers, and saves the populated workbook. | Generate a C# example showing how to populate smart markers from a DataTable using Aspose.Cells WorkbookDesigner, including loading the template and exporting the result file. | Explain step‑by‑step how to bind a DataTable to smart markers in Aspose.Cells, invoke Process, and write the final workbook to disk.
// Common Searches: asp.net core load excel template and fill smart markers from datatable using aspose.cells | c# set custom data source for workbookdesigner smart markers example | how to populate smart markers in an xlsx file with a datatable in Aspose.Cells | using Aspose.Cells WorkbookDesigner to process smart markers from a DataTable
// Tags: Aspose.Cells WorkbookDesigner SetDataSource | C# smart markers DataTable binding | load Excel template with smart markers | process smart markers and save workbook | populate Excel template using DataTable

using System;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    // The sample loads Template.xlsx, creates a DataTable containing product information, assigns the table to WorkbookDesigner as the smart‑marker data source, processes the markers, and saves the filled workbook as Result.xlsx.
    class Program
    {
        static void Main()
        {
            // Load the Excel template that contains smart markers
            Workbook templateWorkbook = new Workbook("Template.xlsx");

            // Create a DataTable that will be used as the custom data source
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("ProductID", typeof(int));
            dataTable.Columns.Add("ProductName", typeof(string));
            dataTable.Columns.Add("Price", typeof(decimal));

            // Populate the DataTable with sample rows
            dataTable.Rows.Add(1, "Laptop", 999.99m);
            dataTable.Rows.Add(2, "Smartphone", 599.49m);
            dataTable.Rows.Add(3, "Tablet", 299.00m);

            // Initialize the WorkbookDesigner and bind the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = templateWorkbook;

            // Assign the DataTable as the data source for smart markers
            designer.SetDataSource(dataTable);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("Result.xlsx");
        }
    }
}
