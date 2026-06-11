using System;
using Aspose.Cells;

class FreezeHeaderDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add header row
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["C1"].PutValue("Score");

        // Add sample data rows
        for (int i = 2; i <= 6; i++)
        {
            worksheet.Cells[i - 1, 0].PutValue(i - 1);                     // ID
            worksheet.Cells[i - 1, 1].PutValue($"User {i - 1}");          // Name
            worksheet.Cells[i - 1, 2].PutValue(50 + (i - 2) * 10);        // Score
        }

        // Freeze the first header row (freeze at cell A2, 1 frozen row, 0 frozen columns)
        worksheet.FreezePanes("A2", 1, 0);

        // Save the workbook
        workbook.Save("FreezeHeaderDemo.xlsx");
    }
}