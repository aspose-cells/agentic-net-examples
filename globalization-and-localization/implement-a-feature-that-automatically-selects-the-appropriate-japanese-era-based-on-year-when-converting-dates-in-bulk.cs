using System;
using Aspose.Cells;
using Aspose.Cells.Settings;

class JapaneseEraBulkConverter
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Set the workbook region to Japan to enable Japanese era formatting
        wb.Settings.Region = CountryCode.Japan;

        // Sample bulk dates placed in column A
        DateTime[] dates = new DateTime[]
        {
            new DateTime(1989, 1, 8),   // Transition from Showa to Heisei
            new DateTime(1990, 5, 15),
            new DateTime(2019, 5, 1),   // Transition from Heisei to Reiwa
            new DateTime(2022, 12, 31)
        };

        for (int i = 0; i < dates.Length; i++)
        {
            ws.Cells[i, 0].PutValue(dates[i]);
        }

        // Create a style that uses Japanese era format
        Style eraStyle = wb.CreateStyle();
        // "ggge" displays era name and year (e.g., "平成31年")
        eraStyle.Custom = "[$-ja-JP]ggge年M月d日";

        // Apply the era style to all cells containing dates
        int lastRow = ws.Cells.MaxDataRow;
        for (int r = 0; r <= lastRow; r++)
        {
            Cell cell = ws.Cells[r, 0];
            if (cell.Type == CellValueType.IsDateTime)
            {
                cell.SetStyle(eraStyle);
            }
        }

        // Save the workbook
        wb.Save("JapaneseEraDates.xlsx");
    }
}