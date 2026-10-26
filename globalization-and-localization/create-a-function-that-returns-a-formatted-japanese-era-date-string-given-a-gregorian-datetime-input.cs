// Title: C# – Convert a Gregorian DateTime to a Japanese era string and write it to an Excel worksheet using Aspose.Cells
// Common Searches: how to format a .NET DateTime as Japanese era (gengō) string | aspnet write japanese era date into Excel cell with Aspose.Cells | c# get japanese era name and year from Gregorian date | convert gregorian calendar date to reiwa era format in C# | save custom formatted japanese era date to .xlsx using Aspose.Cells
// Tags: japanese era representation c# | aspocells write string to excel cell | gregorian to reiwa conversion .net | excel export custom formatted date aspocells | japanese calendar usage with Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example defines a ToJapaneseEraString method that leverages .NET's JapaneseCalendar and Japanese culture info to produce a string such as "令和3年4月30日" from a Gregorian DateTime, handling the first year as "元". It then creates an Aspose.Cells Workbook, writes the formatted era string to cell A1, deletes any existing file with the same name, saves the workbook as JapaneseEra.xlsx, and logs the output path.
public static class JapaneseEraFormatter
{
    /// <param name="date">The Gregorian date to convert.</param>
    /// <returns>A string representing the date in Japanese era format.</returns>
    public static string ToJapaneseEraString(DateTime date)
    {
        // Use the Japanese calendar.
        JapaneseCalendar jpCalendar = new JapaneseCalendar();

        // Obtain the era number for the given date.
        int era = jpCalendar.GetEra(date);

        // Get the era name (e.g., "令和", "平成", "昭和", etc.).
        CultureInfo jaCulture = new CultureInfo("ja-JP");
        string eraName = jaCulture.DateTimeFormat.GetEraName(era);

        // Determine the year within the era.
        int yearInEra = jpCalendar.GetYear(date);
        // In Japanese notation, the first year is expressed as "元" (gan).
        string yearString = yearInEra == 1 ? "元" : yearInEra.ToString();

        // Month and day are straightforward.
        int month = jpCalendar.GetMonth(date);
        int day = jpCalendar.GetDayOfMonth(date);

        // Assemble the formatted string: {Era}{Year}年{Month}月{Day}日
        return $"{eraName}{yearString}年{month}月{day}日";
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Example Gregorian date.
            DateTime gregorian = new DateTime(2021, 4, 30);
            string japaneseEra = JapaneseEraFormatter.ToJapaneseEraString(gregorian);
            Console.WriteLine(japaneseEra); // Output: 令和3年4月30日

            // Save the result into an Excel file using Aspose.Cells.
            string outputPath = "JapaneseEra.xlsx";

            // Ensure we don't attempt to overwrite a locked file.
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            // Create a new workbook and write the formatted date to cell A1.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(japaneseEra);

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
