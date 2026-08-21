// Title: C# – Apply Auto‑Filter with a Between‑Values Custom Filter on an Aspose.Cells Table
// Description: Creates an Excel workbook with a ListObject (table), enables the table's auto‑filter, and uses AutoFilter.Custom with GreaterOrEqual and LessOrEqual criteria to show only rows where the Amount column is between 70 and 130. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# auto filter | Excel table filter | ListObject | AutoFilter.Custom | numeric range filter | FilterOperatorType | sample code | GitHub example
// Common Searches: Aspose.Cells filter rows between two numbers | C# Aspose.Cells custom auto‑filter example | How to use ListObject AutoFilter in Aspose.Cells | Apply numeric range filter to Excel table using Aspose.Cells | Aspose.Cells AutoFilter.Custom between values
// Developer Intent: Enable auto‑filter on a worksheet table and display only the rows whose numeric column values fall within a specified lower‑upper bound.
// Use Cases: Generate a financial statement that lists transactions within a budgeted amount range. | Show a product catalog filtered by price thresholds for quick price‑range analysis. | Export sales data that meets a target interval before further processing.
// AI Prompts: Write C# code with Aspose.Cells to add a ListObject, turn on auto‑filter, and filter a numeric column between 50 and 200. | Explain the parameters of AutoFilter.Custom in Aspose.Cells, especially the isAnd flag for combining two criteria. | Provide a multi‑column filtering example using Aspose.Cells where each column uses a different custom filter.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates an Excel workbook with a ListObject (table), enables the table's auto‑filter, and uses AutoFilter.Custom with GreaterOrEqual and LessOrEqual criteria to show only rows where the Amount column is between 70 and 130. The workbook is saved as an XLSX file.
class TableAutoFilterBetweenValues
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(150);
        sheet.Cells["A5"].PutValue("Date");
        sheet.Cells["B5"].PutValue(60);

        // Define the range of the table (including header)
        int firstRow = 0;   // zero‑based index for row 1
        int firstCol = 0;   // column A
        int lastRow  = 4;   // row 5 (zero‑based)
        int lastCol  = 1;   // column B

        // Add a ListObject (table) covering the data range
        int tableIndex = sheet.ListObjects.Add(firstRow, firstCol, lastRow, lastCol, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Enable auto‑filter for the table
        table.HasAutoFilter = true;

        // Apply a custom filter on the "Amount" column (index 1) to show rows where
        // Amount is between 70 and 130 (inclusive)
        double lowerBound = 70;
        double upperBound = 130;

        // Use the two‑criteria Custom method (AND combination)
        table.AutoFilter.Custom(
            fieldIndex: 1,                     // Amount column
            operatorType1: FilterOperatorType.GreaterOrEqual,
            criteria1: lowerBound,
            isAnd: true,
            operatorType2: FilterOperatorType.LessOrEqual,
            criteria2: upperBound);

        // Refresh the filter to apply the changes
        table.AutoFilter.Refresh();

        // Save the workbook
        workbook.Save("TableAutoFilterBetweenValues.xlsx", SaveFormat.Xlsx);
    }
}
