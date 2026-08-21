// Title: Convert Gregorian Dates to Japanese Era Format with Aspose.Cells (C#)
// Description: Loads a workbook, sets its region to Japan, scans numeric date cells, converts each Excel serial number to a .NET DateTime via CellsHelper.GetDateTimeFromDouble, re‑writes the value, applies the custom format "[$-ja-JP]ggge年M月d日" to show the Japanese era, and saves the file.
// Keywords: Aspose.Cells | Japanese era conversion | Gregorian to Japanese calendar | CellsHelper GetDateTimeFromDouble | workbook.Settings.Region Japan | custom number format ggge | C# Excel date formatting | locale ja-JP | Excel serial date handling
// Common Searches: Aspose.Cells display Japanese era dates | convert Excel serial dates to Japanese calendar C# | preserve date serial while formatting Japanese era | set workbook region to Japan Aspose.Cells | custom number format for Japanese era in Excel
// Developer Intent: Transform every Gregorian date cell in an Excel workbook to a Japanese era representation without altering the underlying serial numbers, using Aspose.Cells.
// Use Cases: Load an existing .xlsx file and configure the workbook for the Japanese locale. | Identify cells that store dates as numeric values (Excel serial numbers). | Use CellsHelper.GetDateTimeFromDouble to obtain a .NET DateTime, then put the value back to retain the original serial. | Apply the number format "[$-ja-JP]ggge年M月d日" so the cell displays the era name and year. | Save the workbook with the updated formatting.
// AI Prompts: Write C# code that reads an Excel workbook, converts all Gregorian date cells to Japanese era format using Aspose.Cells, and saves the result. | Explain step‑by‑step how to keep the original Excel serial number while showing dates in the Japanese calendar with Aspose.Cells. | Create a reusable function that accepts a Workbook object, sets its region to Japan, and formats every date cell with the "[$-ja-JP]ggge年M月d日" pattern.

using System;
using Aspose.Cells;

namespace AsposeCellsJapaneseEraDemo
{
    // Loads a workbook, sets its region to Japan, scans numeric date cells, converts each Excel serial number to a .NET DateTime via CellsHelper.GetDateTimeFromDouble, re‑writes the value, applies the custom format "[$-ja-JP]ggge年M月d日" to show the Japanese era, and saves the file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Set the workbook region to Japan so that Japanese calendar formats are recognized
            workbook.Settings.Region = CountryCode.Japan;

            // Get the first worksheet (adjust if needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Iterate through all used cells
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a numeric value (Excel stores dates as doubles)
                if (cell.IsNumericValue && cell.Type == CellValueType.IsDateTime)
                {
                    // Convert the Excel serial number to a .NET DateTime using the workbook's date system
                    DateTime dt = CellsHelper.GetDateTimeFromDouble(cell.DoubleValue, workbook.Settings.Date1904);

                    // Put the DateTime back into the cell (preserves the original serial value)
                    cell.PutValue(dt);

                    // Apply a custom number format that displays the Japanese era (ggge = era name, year)
                    Style style = cell.GetStyle();
                    style.Custom = "[$-ja-JP]ggge年M月d日";
                    cell.SetStyle(style);
                }
            }

            // Save the modified workbook
            workbook.Save("Output.xlsx");
        }
    }
}
