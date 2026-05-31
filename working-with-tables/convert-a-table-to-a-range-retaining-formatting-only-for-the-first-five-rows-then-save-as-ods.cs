using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;

class ConvertTableToRangeExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (10 rows, 3 columns)
        for (int row = 0; row < 10; row++)
        {
            cells[row, 0].PutValue($"ID {row + 1}");
            cells[row, 1].PutValue($"Name {row + 1}");
            cells[row, 2].PutValue(row * 10);
        }

        // Apply a style to the whole sheet (so formatting exists on all rows)
        Style style = workbook.CreateStyle();
        style.Font.Color = Color.Blue;
        style.Font.IsBold = true;
        sheet.Cells.ApplyStyle(style, new StyleFlag { FontColor = true, FontBold = true });

        // Add a table that covers the data (including header row)
        int tableIndex = sheet.ListObjects.Add(0, 0, 9, 2, true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Convert the table to a range, keeping only the first five rows (0‑4)
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = 4   // zero‑based index; rows 0‑4 correspond to the first five rows
        };
        table.ConvertToRange(options);

        // Save the workbook as ODS
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        workbook.Save("TableConverted.ods", saveOptions);
    }
}