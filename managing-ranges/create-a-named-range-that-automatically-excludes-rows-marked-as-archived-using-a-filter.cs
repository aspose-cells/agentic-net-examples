using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ----- Populate sample data -----
        // Header row
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["C1"].PutValue("Status");

        // Data rows (some marked as "Archived")
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("Item1");
        worksheet.Cells["C2"].PutValue("Active");

        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Item2");
        worksheet.Cells["C3"].PutValue("Archived");

        worksheet.Cells["A4"].PutValue(3);
        worksheet.Cells["B4"].PutValue("Item3");
        worksheet.Cells["C4"].PutValue("Active");

        worksheet.Cells["A5"].PutValue(4);
        worksheet.Cells["B5"].PutValue("Item4");
        worksheet.Cells["C5"].PutValue("Archived");

        // Determine the last row that contains data (zero‑based index)
        int lastDataRowIndex = worksheet.Cells.MaxDataRow; // e.g., 4 for row 5

        // ----- Apply AutoFilter -----
        // Set the filter range to include the header and all data rows
        worksheet.AutoFilter.Range = $"A1:C{lastDataRowIndex + 1}";

        // Show only rows where the Status column (index 2) equals "Active"
        worksheet.AutoFilter.AddFilter(2, "Active");
        worksheet.AutoFilter.Refresh();

        // ----- Create a named range that refers to the data area -----
        // The named range will cover the same area as the table (excluding the header)
        int nameIdx = workbook.Worksheets.Names.Add("ActiveData");
        workbook.Worksheets.Names[nameIdx].RefersTo = $"=Sheet1!$A$2:$C${lastDataRowIndex + 1}";

        // ----- Save the workbook -----
        workbook.Save("NamedRangeExcludingArchived.xlsx");
    }
}