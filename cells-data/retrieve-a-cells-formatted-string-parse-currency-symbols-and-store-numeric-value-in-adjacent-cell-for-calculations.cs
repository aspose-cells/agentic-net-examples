using System;
using System.Globalization;
using Aspose.Cells;

namespace CurrencyParsingExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // lifecycle: create
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: formatted currency strings in column A
            cells["A1"].PutValue("$1,234.56");   // US format
            cells["A2"].PutValue("€2.345,78");   // European format (comma as decimal)
            cells["A3"].PutValue("£3,210");      // No decimal part
            cells["A4"].PutValue("Invalid");     // Non‑numeric string

            // Apply appropriate number formats so that StringValue returns the formatted text
            Style styleUS = cells["A1"].GetStyle();
            styleUS.Number = 164; // Currency format (e.g., $#,##0.00)
            cells["A1"].SetStyle(styleUS);

            Style styleEU = cells["A2"].GetStyle();
            styleEU.Number = 164; // Same format, culture will affect parsing later
            cells["A2"].SetStyle(styleEU);

            Style styleUK = cells["A3"].GetStyle();
            styleUK.Number = 164;
            cells["A3"].SetStyle(styleUK);

            // Define cultures for parsing each cell (could be derived dynamically)
            CultureInfo[] cultures = new CultureInfo[]
            {
                new CultureInfo("en-US"), // for A1
                new CultureInfo("de-DE"), // for A2 (uses comma as decimal)
                new CultureInfo("en-GB"), // for A3
                CultureInfo.InvariantCulture // fallback for A4
            };

            // Process cells in column A (rows 0 to 3)
            for (int row = 0; row <= 3; row++)
            {
                Cell sourceCell = cells[row, 0]; // Column A
                string formattedText = sourceCell.StringValue; // Get formatted string

                // Try to parse using the corresponding culture and Currency style
                double numericValue;
                bool parsed = double.TryParse(
                    formattedText,
                    NumberStyles.Currency,
                    cultures[row],
                    out numericValue);

                // If parsing succeeded, store the numeric value in the adjacent cell (column B)
                if (parsed)
                {
                    Cell targetCell = cells[row, 1]; // Column B
                    targetCell.PutValue(numericValue);
                }
                else
                {
                    // Optionally, write an error indicator
                    cells[row, 1].PutValue("N/A");
                }
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("CurrencyParsingResult.xlsx");
        }
    }
}