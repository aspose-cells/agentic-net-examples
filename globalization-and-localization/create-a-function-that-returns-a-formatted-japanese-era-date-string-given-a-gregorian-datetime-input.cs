// Title: Convert Gregorian DateTime to Japanese Era String with Aspose.Cells for .NET
// Description: A C# utility that creates an Aspose.Cells workbook, sets the region to Japan, writes a Gregorian DateTime to a cell, applies the custom number format "[$-ja-JP]ggge年m月d日", and returns the Japanese era representation (e.g., Reiwa 3年5月1日) while handling errors gracefully.
// Keywords: Aspose.Cells | Japanese era formatting | Gregorian to Japanese era conversion | C# custom number format | Region Japan Aspose.Cells | .NET localization | ggge pattern | Excel date era | Japanese calendar .NET | date formatting Aspose
// Common Searches: Aspose.Cells Japanese era format C# | C# convert Gregorian date to Japanese era string | custom number format ggge Aspose.Cells | set workbook region Japan Aspose | Japanese calendar formatting .NET | format DateTime as Reiwa era using Aspose
// Developer Intent: Create a reusable C# method that receives a Gregorian DateTime and returns the date formatted in the Japanese era using Aspose.Cells.
// Use Cases: Display the current date in Japanese era format in a console or desktop application. | Generate Excel reports that require dates shown in the Japanese calendar. | Return Japanese‑era dates from a web API for Japanese‑language clients. | Batch‑process a list of dates and write era‑formatted strings to a worksheet. | Localize UI elements that show dates to Japanese users.
// AI Prompts: Write unit tests for GetJapaneseEraString covering Meiji, Taisho, Showa, Heisei, and Reiwa eras. | Refactor the function to accept a worksheet, cell address, and optional format string. | Explain each token in the custom format "[$-ja-JP]ggge年m月d日" and how it maps to era, year, month, and day. | Show how to apply the Japanese era format to an entire column of dates in a workbook. | Demonstrate handling of dates earlier than the Meiji era with a fallback format.

using System;
using Aspose.Cells;
using Aspose.Cells.Settings;

namespace Example
{
    // A C# utility that creates an Aspose.Cells workbook, sets the region to Japan, writes a Gregorian DateTime to a cell, applies the custom number format "[$-ja-JP]ggge年m月d日", and returns the Japanese era representation (e.g., Reiwa 3年5月1日) while handling errors gracefully.
    public static class JapaneseEraFormatter
    {
        /// <param name="gregorianDate">The Gregorian date to format.</param>
        /// <returns>Japanese era formatted date string.</returns>
        public static string GetJapaneseEraString(DateTime gregorianDate)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Set regional settings to Japan for Japanese era formatting
                workbook.Settings.Region = CountryCode.Japan;

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Put the Gregorian date into a cell
                Cell cell = sheet.Cells["A1"];
                cell.PutValue(gregorianDate);

                // Apply custom number format for Japanese era
                Style style = cell.GetStyle();
                style.Custom = "[$-ja-JP]ggge年m月d日";
                cell.SetStyle(style);

                // Return the formatted string
                return cell.StringValue;
            }
            catch (Exception ex)
            {
                // Return error information if formatting fails
                return $"Error: {ex.Message}";
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                DateTime now = DateTime.Now;
                string eraString = JapaneseEraFormatter.GetJapaneseEraString(now);
                Console.WriteLine(eraString);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
