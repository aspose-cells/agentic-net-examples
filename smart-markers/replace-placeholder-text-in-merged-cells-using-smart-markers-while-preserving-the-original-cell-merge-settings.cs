// Title: Replace placeholder text in a merged Excel range using Aspose.Cells smart markers while preserving the merge in C#
// AI Prompts: Generate C# code that inserts a smart marker into a merged cell range, binds a DataTable as the data source, runs WorkbookDesigner.Process, and ensures the merged area stays intact. | Show how to use Aspose.Cells WorkbookDesigner to replace a placeholder inside a merged region without breaking the cell merge, including workbook creation and saving.
// Common Searches: how to keep merged cells after processing smart markers with Aspose.Cells C# | replace placeholder in merged Excel cells using Aspose.Cells smart markers | Aspose.Cells WorkbookDesigner preserve merge area when binding DataTable | C# smart marker inside merged range not losing merge | Aspose.Cells merge cells A1:C2 smart marker replacement example
// Tags: smart markers replace placeholder in merged cells | WorkbookDesigner process merged range without breaking merge | Aspose.Cells preserve merged area during smart marker processing | C# bind DataTable to smart marker in merged Excel range | Excel merge cells A1:C2 smart marker example

using System;
using System.Data;
using Aspose.Cells;

// The example creates a workbook, merges cells A1:C2, places a smart marker '&=Employee.Name' inside the merged cell, provides a DataTable with a Name column as the data source, processes the smart marker with WorkbookDesigner, and saves the workbook while the merged area remains unchanged.
class ReplacePlaceholderInMergedCells
{
    static void Main()
    {
        // Create a new workbook that will act as the template
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Merge a range of cells (A1:C2) – this creates a merged area
        // firstRow = 0, firstColumn = 0, totalRows = 2, totalColumns = 3
        cells.Merge(0, 0, 2, 3);

        // Place a smart marker inside the merged cell.
        // The smart marker will be replaced with data from the data source.
        cells["A1"].PutValue("&=Employee.Name");

        // Prepare a simple data source (DataTable) with a column "Name"
        DataTable employeeTable = new DataTable("Employee");
        employeeTable.Columns.Add("Name", typeof(string));
        employeeTable.Rows.Add("Alice Johnson");

        // Initialize WorkbookDesigner, assign the workbook and the data source
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;
        designer.SetDataSource(employeeTable);

        // Process smart markers. The merged area remains intact after processing.
        designer.Process();

        // Save the resulting workbook
        workbook.Save("MergedPlaceholderResult.xlsx");
    }
}
