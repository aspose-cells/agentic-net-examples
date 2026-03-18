using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook that contains the data collection
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the used range of the worksheet (assumes data starts at A1)
        int lastRow = worksheet.Cells.MaxDataRow;      // zero‑based index of the last row with data
        int lastColumn = worksheet.Cells.MaxDataColumn; // zero‑based index of the last column with data

        // Add a ListObject (Excel table) that covers the used range.
        // The 'true' argument indicates that the first row contains headers.
        int tableIndex = worksheet.ListObjects.Add(0, 0, lastRow, lastColumn, true);
        ListObject listObject = worksheet.ListObjects[tableIndex];

        // Optionally set a display name for the table
        listObject.DisplayName = "DataTable";

        // Save the workbook with the newly created ListObject
        workbook.Save("output.xlsx");
    }
}