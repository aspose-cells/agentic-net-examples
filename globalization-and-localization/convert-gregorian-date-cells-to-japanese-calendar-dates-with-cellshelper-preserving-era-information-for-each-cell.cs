using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsJapaneseDateConversion
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (adjust as needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Prepare Japanese culture with JapaneseCalendar
            CultureInfo jpCulture = new CultureInfo("ja-JP");
            jpCulture.DateTimeFormat.Calendar = new JapaneseCalendar();

            // Iterate through all used cells in the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a date value
                    if (cell.Type == CellValueType.IsDateTime)
                    {
                        // Convert the Excel serial number to DateTime using CellsHelper
                        double serial = cell.DoubleValue;
                        DateTime date = CellsHelper.GetDateTimeFromDouble(serial, workbook.Settings.Date1904);

                        // Format the date using Japanese calendar (e.g., "平成31年4月30日")
                        string formatted = date.ToString("gg y年M月d日", jpCulture);

                        // Replace the cell value with the formatted Japanese date string
                        cell.PutValue(formatted);

                        // Ensure the cell is treated as text to preserve the era string
                        Style style = cell.GetStyle();
                        style.Number = 0; // General format (text)
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}