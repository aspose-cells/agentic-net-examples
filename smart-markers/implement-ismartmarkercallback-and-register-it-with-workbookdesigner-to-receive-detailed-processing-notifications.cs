using System;
using System.Data;
using Aspose.Cells;

public class SmartMarkerCallbackDemo : ISmartMarkerCallBack
{
    // This method is called for each smart marker that is processed.
    public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
    {
        Console.WriteLine($"SmartMarker processing - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");
    }

    public static void Main()
    {
        // Initialize WorkbookDesigner
        WorkbookDesigner designer = new WorkbookDesigner();

        // Create a new workbook and assign it to the designer
        designer.Workbook = new Workbook();

        // Add a smart marker to the first worksheet (cell A1)
        Worksheet sheet = designer.Workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("&=$Products.ProductName");

        // Register the callback implementation
        designer.CallBack = new SmartMarkerCallbackDemo();

        // Prepare a DataTable as the data source
        DataTable dt = new DataTable("Products");
        dt.Columns.Add("ProductName", typeof(string));
        dt.Rows.Add("Apple");
        dt.Rows.Add("Banana");
        dt.Rows.Add("Cherry");

        // Bind the data source to the designer
        designer.SetDataSource(dt);

        // Process the smart markers (this will invoke the callback for each marker)
        designer.Process();

        // Save the resulting workbook
        designer.Workbook.Save("SmartMarkerCallbackResult.xlsx");
    }
}