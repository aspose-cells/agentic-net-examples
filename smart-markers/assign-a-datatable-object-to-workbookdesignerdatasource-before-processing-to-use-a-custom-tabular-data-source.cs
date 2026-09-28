// Title: How to bind a DataTable to Aspose.Cells WorkbookDesigner smart markers and generate an .xlsx file in C#
// AI Prompts: Write C# code that creates a DataTable, assigns it to WorkbookDesigner via SetDataSource, processes smart markers, and saves the workbook as an .xlsx file. | Show how to use multiple DataTables with distinct smart‑marker prefixes in Aspose.Cells WorkbookDesigner. | Demonstrate customizing column headers and smart‑marker syntax before binding a DataTable to WorkbookDesigner.
// Common Searches: C# example for using Aspose.Cells WorkbookDesigner with a DataTable as the smart marker source | How to populate Excel rows from a DataTable using Aspose.Cells smart markers | Aspose.Cells SetDataSource method with DataTable and save as xlsx in .NET | Smart markers in Aspose.Cells: binding DataTable to WorkbookDesigner in C#
// Tags: Aspose.Cells WorkbookDesigner SetDataSource DataTable | C# smart markers populate Excel from DataTable | Aspose.Cells generate .xlsx from DataTable | WorkbookDesigner process smart markers C# | DataTable to Excel using Aspose.Cells API

using System;
using System.Data;
using Aspose.Cells;

// The program creates a workbook, defines smart markers, builds a DataTable with product data, assigns the DataTable to WorkbookDesigner via SetDataSource, processes the markers to fill the sheet, and saves the result as DataTableOutput.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define column headers
        sheet.Cells["A1"].PutValue("ProductID");
        sheet.Cells["B1"].PutValue("ProductName");
        sheet.Cells["C1"].PutValue("Price");

        // Place smart markers that will be replaced by DataTable values
        sheet.Cells["A2"].PutValue("&=$ProductID");
        sheet.Cells["B2"].PutValue("&=$ProductName");
        sheet.Cells["C2"].PutValue("&=$Price");

        // Build a DataTable with sample data
        DataTable dt = new DataTable("Products");
        dt.Columns.Add("ProductID", typeof(int));
        dt.Columns.Add("ProductName", typeof(string));
        dt.Columns.Add("Price", typeof(decimal));

        dt.Rows.Add(1, "Laptop", 1200.50m);
        dt.Rows.Add(2, "Smartphone", 799.99m);
        dt.Rows.Add(3, "Tablet", 450.00m);

        // Initialize WorkbookDesigner with the workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Assign the DataTable as the data source for the designer
        designer.SetDataSource(dt);

        // Process the smart markers and populate the worksheet
        designer.Process();

        // Save the populated workbook
        workbook.Save("DataTableOutput.xlsx");
    }
}
