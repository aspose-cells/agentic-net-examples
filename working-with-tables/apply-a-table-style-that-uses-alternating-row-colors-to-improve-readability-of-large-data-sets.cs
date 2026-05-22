using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class ApplyAlternatingRowColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (header + 100 rows)
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Value");
        for (int i = 2; i <= 101; i++)
        {
            sheet.Cells[i - 1, 0].PutValue(i - 1);                     // ID
            sheet.Cells[i - 1, 1].PutValue($"Item {i - 1}");          // Name
            sheet.Cells[i - 1, 2].PutValue((i - 1) * 10);             // Value
        }

        // Convert the range A1:C101 into a table (ListObject)
        int tableIndex = sheet.ListObjects.Add(0, 0, 100, 2, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Apply a built‑in table style that includes row stripes
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Enable alternating row stripe formatting
        table.ShowTableStyleRowStripes = true;

        // Save the workbook with the applied style
        workbook.Save("AlternatingRowColors.xlsx", SaveFormat.Xlsx);
    }
}