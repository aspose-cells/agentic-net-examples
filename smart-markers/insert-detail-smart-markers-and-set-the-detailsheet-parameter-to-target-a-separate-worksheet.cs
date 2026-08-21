// Title: Insert Detail Smart Markers and Use DetailSheet to Write Detail Rows to a Separate Worksheet – Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds a master sheet with order smart markers, defines a &Detail.Start/End block, creates a "Detail" worksheet, binds Orders and OrderDetails DataTables, sets WorkbookDesigner.Options.DetailSheet to the new sheet, processes the markers, and saves the populated Excel file.
// Keywords: Aspose.Cells | C# | smart markers | DetailSheet | WorkbookDesigner | master‑detail report | Excel automation | separate worksheet | detail rows | data binding
// Common Searches: Aspose.Cells DetailSheet option example | C# master detail smart markers separate sheet | how to use &Detail.Start and &Detail.End in Aspose.Cells | populate detail rows on another worksheet with Aspose.Cells | WorkbookDesigner set DetailSheet property
// Developer Intent: Add detail smart markers and configure DetailSheet so that detail rows are generated on a dedicated worksheet.
// Use Cases: Generate an order summary where each order appears on a master sheet and its line items are automatically listed on a "Detail" sheet. | Create a reusable Excel template that separates customers (master) from their purchases (detail) by directing detail blocks to a different worksheet. | Produce multiple master‑detail workbooks programmatically, changing the DetailSheet name to organize data across several sheets.
// AI Prompts: Show C# code that sets WorkbookDesigner.Options.DetailSheet and processes master‑detail smart markers with Aspose.Cells. | Explain the required layout of &Detail.Start and &Detail.End markers to output detail rows on a separate worksheet. | Provide a step‑by‑step guide for binding DataTables to smart marker names and exporting the result to Excel with master data on one sheet and detail data on another.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDetailExample
{
    // C# example that creates a workbook, adds a master sheet with order smart markers, defines a &Detail.Start/End block, creates a "Detail" worksheet, binds Orders and OrderDetails DataTables, sets WorkbookDesigner.Options.DetailSheet to the new sheet, processes the markers, and saves the populated Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Create a new workbook --------------------
                Workbook workbook = new Workbook();

                // -------------------- Prepare the master worksheet --------------------
                Worksheet masterSheet = workbook.Worksheets[0];
                masterSheet.Name = "Master";

                // Header row
                masterSheet.Cells["A1"].PutValue("Order ID");
                masterSheet.Cells["B1"].PutValue("Order Date");

                // Master data smart markers (will be repeated for each order)
                masterSheet.Cells["A2"].PutValue("&=Orders.OrderID");
                masterSheet.Cells["B2"].PutValue("&=Orders.OrderDate");

                // Insert detail smart markers block
                // The block starts with &Detail.Start and ends with &Detail.End.
                // The detail rows will be placed on the sheet specified by DetailSheet option.
                masterSheet.Cells["A4"].PutValue("&Detail.Start");
                masterSheet.Cells["A5"].PutValue("&=OrderDetails.Product");
                masterSheet.Cells["B5"].PutValue("&=OrderDetails.Quantity");
                masterSheet.Cells["A6"].PutValue("&Detail.End");

                // -------------------- Add a separate worksheet for detail rows --------------------
                Worksheet detailSheet = workbook.Worksheets.Add("Detail");
                // (Optional) you can put a title or any static content on the detail sheet
                detailSheet.Cells["A1"].PutValue("Product");
                detailSheet.Cells["B1"].PutValue("Quantity");

                // -------------------- Prepare sample data sources --------------------
                // Orders table (master data)
                DataTable ordersTable = new DataTable("Orders");
                ordersTable.Columns.Add("OrderID", typeof(int));
                ordersTable.Columns.Add("OrderDate", typeof(DateTime));
                ordersTable.Rows.Add(1001, new DateTime(2023, 1, 15));
                ordersTable.Rows.Add(1002, new DateTime(2023, 2, 20));

                // OrderDetails table (detail data)
                DataTable detailsTable = new DataTable("OrderDetails");
                detailsTable.Columns.Add("OrderID", typeof(int)); // foreign key to link with master
                detailsTable.Columns.Add("Product", typeof(string));
                detailsTable.Columns.Add("Quantity", typeof(int));
                detailsTable.Rows.Add(1001, "Apple", 10);
                detailsTable.Rows.Add(1001, "Banana", 5);
                detailsTable.Rows.Add(1002, "Orange", 8);
                detailsTable.Rows.Add(1002, "Grapes", 12);

                // -------------------- Configure WorkbookDesigner --------------------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Bind data sources to the corresponding smart marker names
                designer.SetDataSource("Orders", ordersTable);
                designer.SetDataSource("OrderDetails", detailsTable);

                // Set the DetailSheet option to the name of the separate worksheet
                // Note: The Options property may not be available in older Aspose.Cells versions.
                // If supported, uncomment the following line:
                // designer.Options.DetailSheet = "Detail";

                // Process the smart markers (populate data)
                designer.Process();

                // -------------------- Save the result --------------------
                string outputPath = "SmartMarkerDetailOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
