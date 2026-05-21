using System;
using Aspose.Cells;

namespace AsposeCellsJapaneseEra
{
    public static class JapaneseEraFormatter
    {
        /// <summary>
        /// Returns a formatted Japanese era date string for the specified Gregorian DateTime.
        /// </summary>
        /// <param name="date">Gregorian date to format.</param>
        /// <returns>Japanese era formatted date string.</returns>
        public static string GetJapaneseEraDateString(DateTime date)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Set the workbook region to Japan to ensure Japanese locale handling
            workbook.Settings.Region = CountryCode.Japan;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put the Gregorian date into a cell
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(date);

            // Apply a custom number format that includes the Japanese era (gengō)
            // Format: era name (gg), era year (e), followed by year, month, day in Japanese characters
            Style style = cell.GetStyle();
            style.Custom = "[$-ja-JP]ggge年M月d日";
            cell.SetStyle(style);

            // Retrieve the formatted string value from the cell
            string formattedDate = cell.StringValue;

            // Optionally, clean up resources (save not required for this operation)
            // workbook.Dispose(); // Not necessary as .NET GC will handle it

            return formattedDate;
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            DateTime gregorianDate = new DateTime(2023, 5, 21);
            string japaneseEra = JapaneseEraFormatter.GetJapaneseEraDateString(gregorianDate);
            Console.WriteLine($"Gregorian: {gregorianDate:d} => Japanese Era: {japaneseEra}");
        }
    }
}