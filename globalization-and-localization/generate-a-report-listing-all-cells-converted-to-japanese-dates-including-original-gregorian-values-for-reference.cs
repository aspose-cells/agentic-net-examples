using System;
using Aspose.Cells;
using Aspose.Cells.Settings;

class JapaneseDateReport
{
    static void Main()
    {
        // Create a new workbook (source data)
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Sample Gregorian dates (replace with real data or load from file)
        ws.Cells["A1"].PutValue(new DateTime(2023, 1, 15));
        ws.Cells["A2"].PutValue(new DateTime(2023, 2, 20));
        ws.Cells["B1"].PutValue("Not a date");
        ws.Cells["B2"].PutValue(123);

        // Set regional settings to Japan (affects date handling)
        wb.Settings.Region = CountryCode.Japan;

        // Add a worksheet to hold the report
        int reportIndex = wb.Worksheets.Add();
        Worksheet report = wb.Worksheets[reportIndex];
        report.Name = "JapaneseDateReport";

        // Write header row
        report.Cells[0, 0].PutValue("Cell Address");
        report.Cells[0, 1].PutValue("Gregorian Value");
        report.Cells[0, 2].PutValue("Japanese Date");

        int reportRow = 1;

        // Scan all used cells in the source worksheet
        int maxRow = ws.Cells.MaxDataRow;
        int maxCol = ws.Cells.MaxDataColumn;

        for (int r = 0; r <= maxRow; r++)
        {
            for (int c = 0; c <= maxCol; c++)
            {
                Cell cell = ws.Cells[r, c];

                // Identify cells that contain a DateTime value
                if (cell.Type == CellValueType.IsDateTime)
                {
                    // Original cell address (e.g., "A1")
                    string address = cell.Name;

                    // Original Gregorian date
                    DateTime gregorian = cell.DateTimeValue;

                    // Write address and Gregorian value to the report
                    report.Cells[reportRow, 0].PutValue(address);
                    report.Cells[reportRow, 1].PutValue(gregorian.ToString("yyyy-MM-dd"));

                    // Write the same date with Japanese formatting
                    Cell japCell = report.Cells[reportRow, 2];
                    japCell.PutValue(gregorian);
                    Style style = japCell.GetStyle();
                    // Japanese date format (e.g., "2023年1月15日")
                    style.Custom = "[$-ja-JP]yyyy年m月d日";
                    japCell.SetStyle(style);

                    reportRow++;
                }
            }
        }

        // Save the workbook containing the report
        wb.Save("JapaneseDateReport.xlsx");
    }
}