using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in the range A1:C5
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Score");

        for (int row = 2; row <= 5; row++)
        {
            sheet.Cells[row - 1, 0].PutValue(row - 1);                     // ID
            sheet.Cells[row - 1, 1].PutValue($"User{row - 1}");           // Name
            sheet.Cells[row - 1, 2].PutValue(50 + row * 5);               // Score
        }

        // Convert the populated range into a structured table (ListObject)
        int tableIndex = sheet.ListObjects.Add("A1", "C5", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Configure table properties to leverage advanced features
        table.DisplayName = "UserData";
        table.TableStyleName = "TableStyleMedium9";          // Apply a built‑in style
        table.ShowTotals = true;                             // Enable totals row
        table.ListColumns[2].TotalsCalculation = TotalsCalculation.Average; // Average of Score column

        // Save the workbook with the new structured table
        workbook.Save("StructuredTable.xlsx");
    }
}