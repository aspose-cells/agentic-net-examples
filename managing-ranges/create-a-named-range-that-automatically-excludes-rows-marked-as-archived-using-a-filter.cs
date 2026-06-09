using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeWithFilter
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (headers + rows)
            // Columns: ID | Name | Status
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Status");

            // Row 2
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alpha");
            sheet.Cells["C2"].PutValue("Active");

            // Row 3
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Beta");
            sheet.Cells["C3"].PutValue("Archived");

            // Row 4
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Gamma");
            sheet.Cells["C4"].PutValue("Active");

            // Row 5
            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue("Delta");
            sheet.Cells["C5"].PutValue("Archived");

            // Determine the last data row (zero‑based index)
            int lastDataRow = sheet.Cells.MaxDataRow; // e.g., 4 for rows 0‑4

            // Apply an AutoFilter to the whole data area (including header)
            // SetRange(startRow, startColumn, endRow)
            sheet.AutoFilter.SetRange(0, 0, lastDataRow);

            // Exclude rows where Status = "Archived"
            // Column index for Status is 2 (C)
            // Use a custom filter with NotEqual operator
            sheet.AutoFilter.Custom(2, FilterOperatorType.NotEqual, "Archived");
            sheet.AutoFilter.Refresh();

            // Create a named range that refers to the filtered data area
            // The range includes all rows; the filter hides the "Archived" rows
            int nameIdx = workbook.Worksheets.Names.Add("ActiveData");
            // Build the address string dynamically (e.g., =Sheet1!$A$2:$C$5)
            string refersTo = $"=Sheet1!$A$2:$C${lastDataRow + 1}";
            workbook.Worksheets.Names[nameIdx].RefersTo = refersTo;

            // Save the workbook
            workbook.Save("NamedRangeWithFilter.xlsx");
        }
    }
}