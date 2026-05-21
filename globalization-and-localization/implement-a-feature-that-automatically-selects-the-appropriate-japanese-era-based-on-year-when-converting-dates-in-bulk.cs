using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Settings;

namespace AsposeCellsJapaneseEraDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance (load existing file)
            Workbook workbook = new Workbook("InputData.xlsx");

            // Set workbook region to Japan to ensure Japanese calendar handling
            workbook.Settings.Region = CountryCode.Japan;

            // Get the first worksheet (or iterate all worksheets as needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define Japanese eras with start year (Gregorian) and era name
            var eras = new (int startYear, string name)[]
            {
                (1868, "Meiji"),
                (1912, "Taisho"),
                (1926, "Showa"),
                (1989, "Heisei"),
                (2019, "Reiwa")
            };

            // Iterate through all used cells
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a DateTime value
                if (cell.Type == CellValueType.IsDateTime)
                {
                    DateTime dt = cell.DateTimeValue;
                    string eraName = "";
                    int eraYear = 0;

                    // Find the appropriate era based on the Gregorian year
                    for (int i = eras.Length - 1; i >= 0; i--)
                    {
                        if (dt.Year >= eras[i].startYear)
                        {
                            eraName = eras[i].name;
                            eraYear = dt.Year - eras[i].startYear + 1; // Japanese era year starts at 1
                            break;
                        }
                    }

                    // Fallback if no era matched (should not happen for modern dates)
                    if (string.IsNullOrEmpty(eraName))
                    {
                        eraName = "Unknown";
                        eraYear = dt.Year;
                    }

                    // Build formatted string: e.g., "Reiwa 3年5月21日"
                    string formatted = $"{eraName} {eraYear}年{dt.Month}月{dt.Day}日";

                    // Replace the cell value with the formatted string
                    cell.PutValue(formatted);
                }
            }

            // Save the modified workbook
            workbook.Save("OutputWithJapaneseEra.xlsx");
        }
    }
}