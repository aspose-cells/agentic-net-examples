using System;
using System.Data;
using Aspose.Cells;

class ImportArrayElementByIndexSmartMarker
{
    static void Main()
    {
        // Create a workbook that will act as the template.
        Workbook template = new Workbook();
        Worksheet sheet = template.Worksheets[0];

        // Insert a smart marker that references the second element (index 1) of the "Name" column.
        // Syntax: &=[DataSource].ColumnName[index]
        sheet.Cells["A1"].PutValue("&=Products.Name[1]");

        // Prepare the data source: a DataTable with a single column "Name".
        DataTable products = new DataTable("Products");
        products.Columns.Add("Name", typeof(string));
        products.Rows.Add("Alice");
        products.Rows.Add("Bob");      // This value will be imported (index 1)
        products.Rows.Add("Charlie");

        // Initialize WorkbookDesigner with the template workbook.
        WorkbookDesigner designer = new WorkbookDesigner(template);

        // Bind the DataTable to the smart marker name "Products".
        designer.SetDataSource("Products", products);

        // Process the smart markers – the cell A1 will be replaced with "Bob".
        designer.Process();

        // Save the resulting workbook.
        template.Save("SmartMarkerArrayElementByIndex.xlsx");
    }
}