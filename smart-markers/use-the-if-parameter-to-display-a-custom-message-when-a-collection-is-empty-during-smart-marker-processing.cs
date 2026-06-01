using System;
using System.Data;
using Aspose.Cells;

class SmartMarkerIfDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a smart marker that uses the IF parameter.
        // If the collection "Products" is empty, the text "No products available" will be displayed.
        sheet.Cells["A1"].PutValue("&=$Products.Name?if=IsEmpty?then=No products available");

        // Prepare an empty data source named "Products"
        DataTable productsTable = new DataTable("Products");
        productsTable.Columns.Add("Name", typeof(string));
        // No rows are added – the collection is empty

        // Set up the WorkbookDesigner with the workbook and the empty data source
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };
        designer.SetDataSource(productsTable);

        // Process the smart markers; the custom message will appear because the collection is empty
        designer.Process();

        // Save the resulting workbook
        workbook.Save("SmartMarkerIfDemo.xlsx");
    }
}