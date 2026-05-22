using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods; // Namespace for OdsSaveOptions (if needed)

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (including a header row)
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Score");

        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("John");
        sheet.Cells["C2"].PutValue(85);

        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Mary");
        sheet.Cells["C3"].PutValue(92);

        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["B4"].PutValue("Peter");
        sheet.Cells["C4"].PutValue(78);

        sheet.Cells["A5"].PutValue(4);
        sheet.Cells["B5"].PutValue("Lucy");
        sheet.Cells["C5"].PutValue(88);

        // Add a ListObject (table) that covers the data range A1:C5
        int tableIndex = sheet.ListObjects.Add("A1", "C5", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Apply a built‑in table style (optional, demonstrates formatting preservation)
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Convert the table to a normal range, preserving formatting up to the last data row.
        // The last data row index is 4 (zero‑based, corresponds to row 5 in Excel).
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = 4
        };
        table.ConvertToRange(options);

        // Save the workbook as an ODS file using default OdsSaveOptions.
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        workbook.Save("ConvertedTable.ods", saveOptions);
    }
}