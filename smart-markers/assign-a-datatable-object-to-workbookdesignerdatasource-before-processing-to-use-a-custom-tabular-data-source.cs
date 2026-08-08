// Title: C# – Assign a DataTable to WorkbookDesigner for Smart Marker Processing in Aspose.Cells
// Description: Demonstrates how to create a workbook template, add smart markers, build a custom DataTable, bind it to WorkbookDesigner with SetDataSource, process the markers, and save the populated Excel file.
// Keywords: Aspose.Cells | WorkbookDesigner | DataTable binding | smart markers .NET | C# Excel template | SetDataSource overload | populate Excel from DataTable
// Common Searches: bind DataTable to WorkbookDesigner Aspose.Cells | smart markers using custom DataTable C# | set data source for WorkbookDesigner | Aspose.Cells populate template from DataTable
// Developer Intent: Bind a custom DataTable to WorkbookDesigner and execute smart marker processing to generate a filled Excel workbook.
// Use Cases: Generate product catalogs by merging a DataTable of items into an Excel template. | Create financial statements from database query results using smart markers. | Automate report generation where any DataTable can be injected into a pre‑designed worksheet.
// AI Prompts: Show C# code that assigns a DataTable to WorkbookDesigner and processes smart markers in Aspose.Cells. | Explain how to use multiple DataTables with WorkbookDesigner SetDataSource overloads. | What steps are needed to handle column name mismatches between a DataTable and smart markers in Aspose.Cells?

using System;
using System.Data;
using Aspose.Cells;

// Demonstrates how to create a workbook template, add smart markers, build a custom DataTable, bind it to WorkbookDesigner with SetDataSource, process the markers, and save the populated Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook (template)
        Workbook workbook = new Workbook();

        // Add smart markers to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("ProductID");      // Header
        sheet.Cells["B1"].PutValue("ProductName");    // Header
        sheet.Cells["A2"].PutValue("&=$ProductID");   // Smart marker for ProductID
        sheet.Cells["B2"].PutValue("&=$ProductName"); // Smart marker for ProductName

        // Build a custom DataTable as the data source
        DataTable dataTable = new DataTable("Products");
        dataTable.Columns.Add("ProductID", typeof(int));
        dataTable.Columns.Add("ProductName", typeof(string));
        dataTable.Rows.Add(1, "Laptop");
        dataTable.Rows.Add(2, "Smartphone");
        dataTable.Rows.Add(3, "Tablet");

        // Initialize WorkbookDesigner and bind the workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Assign the DataTable to the designer (SetDataSource overload)
        designer.SetDataSource(dataTable);

        // Process the smart markers using the assigned data source
        designer.Process();

        // Save the resulting workbook
        workbook.Save("CustomDataTableOutput.xlsx");
    }
}
