using Aspose.Cells;
using System;
using System.Data;

public class SmartMarkerNotifyDemo : ISmartMarkerCallBack
{
    // This method is called for each smart marker row insertion
    public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
    {
        Console.WriteLine($"Row inserted - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");
    }

    public static void Run()
    {
        // Create a new workbook and place smart markers in the first worksheet
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("&=$Employees.Name");
        ws.Cells["B1"].PutValue("&=$Employees.Age");

        // Initialize WorkbookDesigner and assign the callback to receive notifications
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook,
            CallBack = new SmartMarkerNotifyDemo()
        };

        // Prepare a DataTable as the data source for the smart markers
        DataTable dt = new DataTable("Employees");
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Age", typeof(int));
        dt.Rows.Add("Alice", 30);
        dt.Rows.Add("Bob", 25);
        dt.Rows.Add("Charlie", 28);

        // Set the data source and process the smart markers
        designer.SetDataSource(dt);
        designer.Process(true); // true = preserve unrecognized markers (optional)

        // Save the resulting workbook
        workbook.Save("SmartMarkerNotifyDemo.xlsx");
    }
}

class Program
{
    static void Main()
    {
        SmartMarkerNotifyDemo.Run();
    }
}