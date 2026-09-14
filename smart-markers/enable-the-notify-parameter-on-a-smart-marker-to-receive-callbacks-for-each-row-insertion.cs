// Title: Enable notification callbacks for each smart‑marker row insertion in Aspose.Cells for .NET (C#)
// AI Prompts: Create a class that implements the smart‑marker callback interface and assign it to WorkbookDesigner.CallBack to receive a callback whenever a smart‑marker row is inserted. | Demonstrate how to turn on notifications for a smart‑marker range and process the data source so that each inserted row is logged during WorkbookDesigner.Process.
// Common Searches: Aspose.Cells C# enable smart marker row insertion callback | How to attach a smart‑marker callback to WorkbookDesigner in .NET | Receive per‑row notifications when processing smart markers in Aspose.Cells | Configure smart marker notifications for a DataTable source in C# | Log each inserted row from smart markers using Aspose.Cells API
// Tags: ISmartMarkerCallBack C# example | activate notify parameter Aspose.Cells | smart marker callback processing | WorkbookDesigner row insertion notification | Aspose.Cells smart marker event handling

using System;
using System.Data;
using Aspose.Cells;

// The example creates a workbook with smart markers, defines a DataTable as the data source, implements a custom smart‑marker callback class, assigns it to WorkbookDesigner.CallBack, processes the smart markers so the callback logs each inserted row, and saves the resulting Excel file.
public class SmartMarkerNotifyDemo : ISmartMarkerCallBack
{
    // This method will be called for each smart marker row insertion
    public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
    {
        Console.WriteLine($"Inserted row - Sheet:{sheetIndex}, Row:{rowIndex}, Column:{colIndex}, Table:{tableName}, Column:{columnName}");
    }

    public static void Main()
    {
        // Create a new workbook and add smart markers
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("&=Table1.Column1");
        sheet.Cells["B1"].PutValue("&=Table1.Column2");
        // Mark the range that contains smart markers (required when using callbacks)
        sheet.Cells.CreateRange("A1:B1").Name = "_CellsSmartMarkers";

        // Prepare a data source (DataTable) for the smart markers
        DataTable table = new DataTable("Table1");
        table.Columns.Add("Column1", typeof(string));
        table.Columns.Add("Column2", typeof(int));
        table.Rows.Add("First", 100);
        table.Rows.Add("Second", 200);
        table.Rows.Add("Third", 300);

        // Initialize WorkbookDesigner and assign the workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Enable notification callbacks by assigning an implementation of ISmartMarkerCallBack
        designer.CallBack = new SmartMarkerNotifyDemo();

        // Set the data source for the smart markers
        designer.SetDataSource(table);

        // Process the smart markers; the Process method of the callback will be invoked for each inserted row
        designer.Process(true);

        // Save the resulting workbook
        designer.Workbook.Save("SmartMarkerNotifyDemo.xlsx");
    }
}
