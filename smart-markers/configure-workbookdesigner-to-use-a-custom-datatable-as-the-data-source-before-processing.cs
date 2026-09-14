// Title: How to bind a custom DataTable to WorkbookDesigner and generate Excel with smart markers using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a DataTable, defines smart markers in a workbook, sets the DataTable as the WorkbookDesigner data source, processes the markers, and saves the result as an XLSX file. | Show the step‑by‑step configuration of WorkbookDesigner with a DataTable named "Products" to export populated data to Excel via smart markers.
// Common Searches: Aspose.Cells C# set DataTable as data source for WorkbookDesigner smart markers | example of using WorkbookDesigner with a custom DataTable to fill an Excel template | how to populate Excel rows from a DataTable using smart markers in Aspose.Cells .NET
// Tags: WorkbookDesigner SetDataSource DataTable | Aspose.Cells smart markers with DataTable | populate Excel from DataTable using WorkbookDesigner | C# Aspose.Cells custom data source for smart markers

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to configure WorkbookDesigner with a custom DataTable as the data source.
    // The sample creates a DataTable called "Products", adds sample rows, inserts smart marker syntax into a new workbook, assigns the workbook to a WorkbookDesigner, sets the DataTable as the data source via SetDataSource, processes the markers to fill the sheet, and saves the workbook as an XLSX file.
    public class WorkbookDesignerWithDataTable
    {
        public static void Main()
        {
            // 1. Create a sample DataTable that will serve as the custom data source.
            DataTable dataTable = new DataTable("Products");
            dataTable.Columns.Add("ProductID", typeof(int));
            dataTable.Columns.Add("ProductName", typeof(string));
            dataTable.Columns.Add("Price", typeof(decimal));

            // Add some rows.
            dataTable.Rows.Add(1, "Laptop", 1200.00m);
            dataTable.Rows.Add(2, "Smartphone", 799.99m);
            dataTable.Rows.Add(3, "Tablet", 450.50m);

            // 2. Create a new workbook (or load an existing template if needed).
            Workbook workbook = new Workbook();

            // 3. Insert smart markers into the worksheet where data should be populated.
            //    The marker syntax "&=$Products.ColumnName" binds to the DataTable named "Products".
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("&=$Products.ProductID");
            sheet.Cells["B1"].PutValue("&=$Products.ProductName");
            sheet.Cells["C1"].PutValue("&=$Products.Price");

            // 4. Initialize WorkbookDesigner and assign the workbook.
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;

            // 5. Set the custom DataTable as the data source.
            //    This uses the SetDataSource(DataTable) overload as defined in the API.
            designer.SetDataSource(dataTable);

            // 6. Process the smart markers – this populates the worksheet with the DataTable data.
            designer.Process();

            // 7. Save the resulting workbook.
            workbook.Save("WorkbookDesigner_With_DataTable.xlsx");
        }
    }
}
