using System;
using Aspose.Cells;

namespace AsposeCellsJapaneseEraDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Ensure the workbook uses the Japanese regional settings
            workbook.Settings.Region = CountryCode.Japan;

            // Get the first worksheet (adjust if needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Iterate through all used cells
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a date (Excel stores dates as numbers)
                    if (cell.Type == CellValueType.IsDateTime || 
                        (cell.IsNumericValue && IsDateFormatted(cell)))
                    {
                        // Convert the Excel serial number to a .NET DateTime
                        double serial = cell.DoubleValue;
                        DateTime date = CellsHelper.GetDateTimeFromDouble(serial, workbook.Settings.Date1904);

                        // Put the DateTime back into the cell
                        cell.PutValue(date);

                        // Apply Japanese era format (e.g., "平成31年4月30日")
                        Style style = cell.GetStyle();
                        style.Custom = "[$-ja-JP]ggge年m月d日";
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }

        // Helper method to determine if a numeric cell is formatted as a date
        private static bool IsDateFormatted(Cell cell)
        {
            // Retrieve the number format index; 14-22 are built‑in date formats in Excel
            int number = cell.GetStyle().Number;
            if (number >= 14 && number <= 22) return true;

            // Check for custom date format containing date placeholders
            string custom = cell.GetStyle().Custom;
            if (!string.IsNullOrEmpty(custom) &&
                (custom.Contains("y") || custom.Contains("m") || custom.Contains("d")))
                return true;

            return false;
        }
    }
}