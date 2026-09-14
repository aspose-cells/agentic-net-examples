// Title: Implement ISmartMarkerCallBack in Aspose.Cells for .NET to modify cell values after smart marker processing
// AI Prompts: Write a C# ISmartMarkerCallBack that prefixes string cells with "Modified-" and multiplies numeric cells by 10 during smart marker processing. | Show how to assign the callback to WorkbookDesigner, bind a DataTable, invoke Process(), and save the workbook with the transformed values.
// Common Searches: how to apply a custom callback to change smart marker output in Aspose.Cells C# | example of scaling numeric values in smart marker callback Aspose.Cells .NET | prefix text to string cells using ISmartMarkerCallBack after processing | registering ISmartMarkerCallBack with WorkbookDesigner for data table | modify smart marker generated cells before saving workbook Aspose.Cells
// Tags: ISmartMarkerCallBack cell value transformation Aspose.Cells | prepend string in smart marker callback C# | multiply numeric cells in smart marker processing | register callback with WorkbookDesigner .NET | smart marker post‑processing modification example | Aspose.Cells callback for DataTable binding

using System;
using System.Data;
using Aspose.Cells;

// The example defines a SmartMarkerCallbackDemo that implements ISmartMarkerCallBack. During the Process() call, each populated smart‑marker cell is accessed via a static workbook context; string cells are prefixed with "Modified-" and numeric cells are multiplied by 10. The callback is attached to WorkbookDesigner, a DataTable is bound, processing is executed, and the resulting workbook is saved.
public class SmartMarkerCallbackDemo : ISmartMarkerCallBack
{
    // This method is called for each smart marker cell after it has been populated.
    public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
    {
        // Retrieve the workbook that is being processed.
        Workbook wb = CallbackContext.CurrentWorkbook;
        if (wb == null) return;

        Worksheet ws = wb.Worksheets[sheetIndex];
        Cell cell = ws.Cells[rowIndex, colIndex];

        // Example modification:
        // - If the cell contains a string, prepend "Modified-".
        // - If the cell contains a numeric value, multiply it by 10.
        if (cell.Type == CellValueType.IsString)
        {
            cell.PutValue("Modified-" + cell.StringValue);
        }
        else if (cell.Type == CellValueType.IsNumeric)
        {
            cell.PutValue(cell.DoubleValue * 10);
        }
    }
}

// Simple static holder to give the callback access to the workbook instance.
public static class CallbackContext
{
    public static Workbook CurrentWorkbook { get; set; }
}

public class Program
{
    public static void Main()
    {
        // 1. Create a workbook with smart markers.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("&=$Table.Column1"); // String column
        sheet.Cells["B1"].PutValue("&=$Table.Column2"); // Numeric column

        // 2. Prepare a data source.
        DataTable dt = new DataTable("Table");
        dt.Columns.Add("Column1", typeof(string));
        dt.Columns.Add("Column2", typeof(double));
        dt.Rows.Add("Item1", 5);
        dt.Rows.Add("Item2", 10);

        // 3. Store the workbook reference for the callback.
        CallbackContext.CurrentWorkbook = workbook;

        // 4. Set up the WorkbookDesigner with the callback.
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook,
            CallBack = new SmartMarkerCallbackDemo()
        };

        // 5. Bind the data source and process smart markers.
        designer.SetDataSource(dt);
        designer.Process(); // Callback runs during this call.

        // 6. Save the workbook after processing and callback modifications.
        workbook.Save("SmartMarkerCallbackResult.xlsx");
    }
}
