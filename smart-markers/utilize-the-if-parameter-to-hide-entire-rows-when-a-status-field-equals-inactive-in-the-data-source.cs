using System;
using Aspose.Cells;

class HideInactiveRows
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample data with a header row
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Status");
        cells["A2"].PutValue(1);
        cells["B2"].PutValue("Active");
        cells["A3"].PutValue(2);
        cells["B3"].PutValue("Inactive");
        cells["A4"].PutValue(3);
        cells["B4"].PutValue("Active");
        cells["A5"].PutValue(4);
        cells["B5"].PutValue("Inactive");

        // Column index for the Status field (B column => index 1)
        int statusColumnIndex = 1;

        // Iterate through all data rows (skip header at row 0)
        for (int row = 1; row <= worksheet.Cells.MaxDataRow; row++)
        {
            // Hide the row if the Status cell equals "Inactive"
            if (worksheet.Cells[row, statusColumnIndex].StringValue == "Inactive")
            {
                worksheet.Cells.HideRow(row);
            }
        }

        // Save the workbook with hidden rows
        workbook.Save("HiddenInactiveRows.xlsx", SaveFormat.Xlsx);
    }
}