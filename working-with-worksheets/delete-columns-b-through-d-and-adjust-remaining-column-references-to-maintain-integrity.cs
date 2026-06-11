using System;
using Aspose.Cells;

class DeleteColumnsExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in columns A to E
        cells["A1"].PutValue("Col A");
        cells["B1"].PutValue("Col B");
        cells["C1"].PutValue("Col C");
        cells["D1"].PutValue("Col D");
        cells["E1"].PutValue("Col E");
        cells["A2"].PutValue(10);
        cells["B2"].PutValue(20);
        cells["C2"].PutValue(30);
        cells["D2"].PutValue(40);
        cells["E2"].PutValue(50);

        // Add a formula that references the range A2:E2
        cells["F1"].Formula = "=SUM(A2:E2)";

        // Delete columns B through D (indexes 1 to 3) and update references
        cells.DeleteColumns(1, 3, true);

        // After deletion, the formula should be adjusted automatically
        Console.WriteLine("Updated formula in F1: " + cells["F1"].Formula);

        // Save the modified workbook
        workbook.Save("DeletedColumns.xlsx");
    }
}