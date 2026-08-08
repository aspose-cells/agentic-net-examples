// Title: C# – Custom ISmartMarkerCallBack in Aspose.Cells to Alter Cell Values After Smart‑Marker Merge
// Description: Shows how to implement ISmartMarkerCallBack for Aspose.Cells (.NET), attach it to WorkbookDesigner, and modify each cell (e.g., prepend a prefix) right after smart‑marker processing and before the workbook is saved.
// Keywords: Aspose.Cells | C# | ISmartMarkerCallBack | smart marker callback | modify cell after merge | WorkbookDesigner | custom callback | data merge | Excel automation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells custom smart marker callback C# | How to change cell value after smart marker processing | ISmartMarkerCallBack example | Modify smart marker output before saving workbook | Aspose.Cells .NET data merge callback
// Developer Intent: Create a callback that receives sheet, row, column indices and table/column names during smart‑marker processing and updates the cell content prior to saving the workbook.
// Use Cases: Add an audit prefix to every merged value. | Apply conditional text or formatting based on source column data. | Insert status flags or timestamps after data merge. | Replace placeholders with localized strings during export.
// AI Prompts: Write C# code that implements ISmartMarkerCallBack to prepend "Modified_" to each smart‑marker result and integrates it with WorkbookDesigner. | Provide a step‑by‑step guide for setting up and registering a custom smart‑marker callback in Aspose.Cells for .NET, including data source binding and workbook saving. | Show how to access sheet, row, and column indices inside the Process method to change cell value and apply formatting after smart‑marker processing.

using System;
using System.Data;
using Aspose.Cells;

// Shows how to implement ISmartMarkerCallBack for Aspose.Cells (.NET), attach it to WorkbookDesigner, and modify each cell (e.g., prepend a prefix) right after smart‑marker processing and before the workbook is saved.
public class MySmartMarkerCallback : ISmartMarkerCallBack
{
    private readonly Workbook _workbook;

    public MySmartMarkerCallback(Workbook workbook)
    {
        _workbook = workbook;
    }

    // This method is called for each smart marker after it has been processed.
    // Here we change the cell value to indicate that the callback was executed.
    public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
    {
        Worksheet sheet = _workbook.Worksheets[sheetIndex];
        Cell cell = sheet.Cells[rowIndex, colIndex];
        cell.PutValue($"Modified_{tableName}_{columnName}");
    }
}

public class SmartMarkerCallbackDemo
{
    public static void Run()
    {
        // Create a new workbook and place a smart marker in cell A1.
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("&=Table1.Column1"); // Smart marker syntax

        // Prepare a simple data source.
        DataTable dt = new DataTable("Table1");
        dt.Columns.Add("Column1", typeof(string));
        dt.Rows.Add("OriginalValue");

        // Set up the WorkbookDesigner, assign the data source and the callback.
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = wb,
            CallBack = new MySmartMarkerCallback(wb)
        };
        designer.SetDataSource(dt);

        // Process the smart markers. The callback will modify the cell value.
        designer.Process(true);

        // Save the resulting workbook.
        wb.Save("SmartMarkerCallbackResult.xlsx");
    }
}

// Entry point for demonstration.
class Program
{
    static void Main()
    {
        SmartMarkerCallbackDemo.Run();
        Console.WriteLine("Workbook saved as SmartMarkerCallbackResult.xlsx");
    }
}
