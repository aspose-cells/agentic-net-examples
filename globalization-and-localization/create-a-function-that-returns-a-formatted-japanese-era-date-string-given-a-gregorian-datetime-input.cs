// Title: C# – Convert Gregorian DateTime to Japanese Era String with Aspose.Cells
// Description: A static C# method that creates an in‑memory Workbook, sets its region to Japan, inserts a Gregorian DateTime into a cell, applies the locale‑specific format "[$-ja-JP]ggge年M月d日", and returns the resulting Japanese era representation (e.g., "令和3年5月10日") without any file I/O.
// Keywords: Aspose.Cells | Japanese era formatting | ggge format | C# date conversion | Japan locale | in‑memory workbook | .NET date formatting | Gregorian to Reiwa | custom number format
// Common Searches: Aspose.Cells format date as Japanese era C# | ggge custom number format example | convert Gregorian date to Reiwa string .NET | set workbook region to Japan Aspose.Cells | Japanese era date string without saving Excel
// Developer Intent: Provide a reusable function that returns a Japanese era formatted date string from a Gregorian DateTime using Aspose.Cells.
// Use Cases: Create Excel reports for Japanese clients where dates appear in the era style (令和, 平成, etc.). | Localize invoice or receipt dates before exporting spreadsheets to the Japanese market. | Generate multilingual workbooks that automatically display era names based on the workbook's region setting.
// AI Prompts: Write a C# function that uses Aspose.Cells to format a DateTime as a Japanese era string without writing a file. | Show how to configure a workbook's region to Japan and apply the "[$-ja-JP]ggge年M月d日" format to a cell. | Explain how to extract the formatted era string from a cell's StringValue after applying the locale‑specific format.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A static C# method that creates an in‑memory Workbook, sets its region to Japan, inserts a Gregorian DateTime into a cell, applies the locale‑specific format "[$-ja-JP]ggge年M月d日", and returns the resulting Japanese era representation (e.g., "令和3年5月10日") without any file I/O.
    public static class JapaneseEraFormatter
    {
        /// <param name="date">Gregorian date to be converted.</param>
        /// <returns>Japanese era formatted string, e.g., "令和3年5月10日".</returns>
        public static string GetJapaneseEraString(DateTime date)
        {
            // Create a new workbook (in-memory, no file I/O required)
            Workbook workbook = new Workbook();

            // Set the workbook region to Japan to ensure Japanese locale is used
            workbook.Settings.Region = CountryCode.Japan;

            // Use the first worksheet and a single cell to leverage Aspose.Cells formatting engine
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Put the Gregorian DateTime value into the cell
            cell.PutValue(date);

            // Apply a custom number format that includes the Japanese era (ggge)
            // The format string "[$-ja-JP]ggge年M月d日" tells Aspose.Cells to use the Japanese
            // locale and display the era name (e.g., "令和") followed by the year, month, and day.
            Style style = cell.GetStyle();
            style.Custom = "[$-ja-JP]ggge年M月d日";
            cell.SetStyle(style);

            // Retrieve the formatted string representation from the cell
            return cell.StringValue;
        }

        // Example usage
        public static void Run()
        {
            DateTime gregorianDate = new DateTime(2023, 5, 10);
            string japaneseEra = GetJapaneseEraString(gregorianDate);
            Console.WriteLine($"Gregorian: {gregorianDate:d} => Japanese Era: {japaneseEra}");
        }
    }

    // Entry point for the application
    public static class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                JapaneseEraFormatter.Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
