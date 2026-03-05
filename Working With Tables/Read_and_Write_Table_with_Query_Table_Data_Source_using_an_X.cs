using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Tables;

class QueryTableExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // 1. Populate sample data that will be used as a table
        // -------------------------------------------------
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Score");

        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["C2"].PutValue(85);

        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");
        sheet.Cells["C3"].PutValue(92);

        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["B4"].PutValue("Charlie");
        sheet.Cells["C4"].PutValue(78);

        // -------------------------------------------------
        // 2. Create a ListObject (Excel table)
        // -------------------------------------------------
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
        int listObjectIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
        ListObject listObject = sheet.ListObjects[listObjectIndex];
        listObject.DisplayName = "StudentScores";

        // -------------------------------------------------
        // 3. Export the table data to a DataTable (read operation)
        // -------------------------------------------------
        // Export the range that contains the table (including header row)
        DataTable exportedTable = sheet.Cells.ExportDataTable(0, 0, 5, 3, true);

        Console.WriteLine("Exported DataTable contents:");
        foreach (DataRow row in exportedTable.Rows)
        {
            Console.WriteLine($"{row["ID"]}, {row["Name"]}, {row["Score"]}");
        }

        // -------------------------------------------------
        // 4. Modify the DataTable (e.g., add a new row)
        // -------------------------------------------------
        DataRow newRow = exportedTable.NewRow();
        newRow["ID"] = 4;
        newRow["Name"] = "Diana";
        newRow["Score"] = 88;
        exportedTable.Rows.Add(newRow);

        // -------------------------------------------------
        // 5. Import the modified DataTable back into the worksheet (write operation)
        // -------------------------------------------------
        // Clear existing data first to avoid duplication
        sheet.Cells.Clear();

        // Import data starting at cell A1, showing field names as headers
        ImportTableOptions importOptions = new ImportTableOptions
        {
            IsFieldNameShown = true
        };
        sheet.Cells.ImportData(exportedTable, 0, 0, importOptions);

        // -------------------------------------------------
        // 6. (Optional) Demonstrate reading a QueryTable if one exists
        // -------------------------------------------------
        if (sheet.QueryTables.Count > 0)
        {
            QueryTable qt = sheet.QueryTables[0];
            Console.WriteLine($"QueryTable Name: {qt.Name}");
            Console.WriteLine($"Result Range: {qt.ResultRange.Address}");
        }
        else
        {
            Console.WriteLine("No QueryTable objects found in the worksheet.");
        }

        // -------------------------------------------------
        // 7. Save the workbook to an XLSX file
        // -------------------------------------------------
        workbook.Save("QueryTableDemo.xlsx");
        Console.WriteLine("Workbook saved as QueryTableDemo.xlsx");
    }
}