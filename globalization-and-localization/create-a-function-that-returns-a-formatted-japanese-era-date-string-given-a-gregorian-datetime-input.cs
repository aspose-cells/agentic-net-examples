// Title: How to format a Gregorian DateTime as a Japanese era (wareki) string in C# using CultureInfo and JapaneseCalendar
// Common Searches: C# format DateTime to Japanese era string using JapaneseCalendar | Convert Gregorian date to wareki format with .NET CultureInfo | How to display Japanese era (gg) and year (yy) in C# date formatting | Create custom date format ggyy年M月d日 in C# | Example of using ja-JP CultureInfo for Japanese era dates in .NET
// Tags: C# JapaneseCalendar date formatting | Japanese era wareki string conversion .NET | custom ggyy年M月d日 format C# | ja-JP CultureInfo for era dates | Gregorian to Japanese era conversion C#

using System;
using System.Globalization;

// The ToJapaneseEraString method builds a Japanese CultureInfo with a JapaneseCalendar, defines the custom format string "ggyy年M月d日", and returns the supplied Gregorian DateTime formatted in Japanese era notation (e.g., "令和3年5月1日").
public static class JapaneseEraFormatter
{
    /// <param name="date">The Gregorian date to convert.</param>
    /// <returns>A string representing the date in Japanese era format.</returns>
    public static string ToJapaneseEraString(DateTime date)
    {
        // Create a CultureInfo for Japanese (Japan) and set its calendar to JapaneseCalendar.
        CultureInfo japaneseCulture = new CultureInfo("ja-JP", false);
        japaneseCulture.DateTimeFormat.Calendar = new JapaneseCalendar();

        // Define the format:
        // gg  - era name (e.g., "令和")
        // yy  - year of the era (two‑digit; use yyyy for four‑digit)
        // M   - month number without leading zero
        // d   - day number without leading zero
        // The literal characters "年", "月", "日" are added for Japanese date notation.
        const string format = "ggyy年M月d日";

        // Use the custom culture to format the date.
        return date.ToString(format, japaneseCulture);
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Example usage of the JapaneseEraFormatter.
            DateTime gregorian = new DateTime(2021, 5, 1);
            string wareki = JapaneseEraFormatter.ToJapaneseEraString(gregorian);
            Console.WriteLine(wareki); // Expected output: 令和3年5月1日
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
