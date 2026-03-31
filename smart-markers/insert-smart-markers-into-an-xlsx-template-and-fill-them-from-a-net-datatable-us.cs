using System;
using System.Data;
using Aspose.Cells;

class SmartMarkerExample
{
    static void Main()
    {
        // Create a workbook and add a worksheet with smart markers
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Products";

        // Add headers
        sheet.Cells["A1"].PutValue("Product ID");
        sheet.Cells["B1"].PutValue("Product Name");
        sheet.Cells["C1"].PutValue("Price");

        // Add smart markers (they start from row 2)
        sheet.Cells["A2"].PutValue("&Products.ProductID");
        sheet.Cells["B2"].PutValue("&Products.ProductName");
        sheet.Cells["C2"].PutValue("&Products.Price");

        // Create a WorkbookDesigner and assign the workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Build a DataTable that matches the smart marker table name ("Products")
        DataTable dt = new DataTable("Products");
        dt.Columns.Add("ProductID", typeof(int));
        dt.Columns.Add("ProductName", typeof(string));
        dt.Columns.Add("Price", typeof(double));

        // Add sample rows
        dt.Rows.Add(1, "Laptop", 1200.50);
        dt.Rows.Add(2, "Smartphone", 799.99);
        dt.Rows.Add(3, "Tablet", 450.00);

        // Set the DataTable as the data source for the designer
        designer.SetDataSource(dt);

        // Process the smart markers – this fills the worksheet with the DataTable data
        designer.Process();

        // Save the populated workbook
        string outputPath = "SmartMarkerResult.xlsx";
        designer.Workbook.Save(outputPath);
    }
}