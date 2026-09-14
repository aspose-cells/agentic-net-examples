// Title: Convert a Gregorian DateTime to a Japanese era string and save it in an Excel workbook using Aspose.Cells (C#)
// AI Prompts: Write a C# method that takes a DateTime and returns a Japanese era formatted string (e.g., "R3年5月1日"), correctly handling era start dates and the "元" first year. | Create an Aspose.Cells workbook, place the formatted Japanese era string into cell A1, and save the file as an .xlsx document. | Add validation to throw an exception for dates earlier than the Meiji era and ensure the output directory exists before saving.
// Common Searches: c# how to format date as Japanese era using Aspose.Cells | convert DateTime to Reiwa era string in .NET | write localized Japanese calendar date to Excel with Aspose.Cells | example of Japanese era date conversion for Excel export in C# | Aspose.Cells save workbook with custom Japanese era date format
// Tags: c# japanese era date conversion | aspocells add era formatted text | japanese calendar formatting .net | excel workbook save custom date string | era start date handling c#

using System;
using System.IO;
using Aspose.Cells;

namespace JapaneseEraExample
{
    // Utility class to convert Gregorian dates to Japanese era format.
    // The example defines a JapaneseEraFormatter with a ToJapaneseEraString method that converts a Gregorian DateTime to a Japanese era representation (e.g., "R3年5月1日"), then creates an Aspose.Cells Workbook, writes the formatted string to cell A1, and saves the workbook as JapaneseEraDate.xlsx.
    public static class JapaneseEraFormatter
    {
        // Represents a Japanese era with its name, abbreviation and start date.
        private class Era
        {
            public string Name { get; }
            public string Abbreviation { get; }
            public DateTime StartDate { get; }

            public Era(string name, string abbreviation, DateTime startDate)
            {
                Name = name;
                Abbreviation = abbreviation;
                StartDate = startDate;
            }
        }

        // List of eras in chronological order (oldest first).
        private static readonly Era[] Eras = new Era[]
        {
            new Era("Meiji",   "M", new DateTime(1868,  1, 25)),
            new Era("Taisho",  "T", new DateTime(1912,  7, 30)),
            new Era("Showa",   "S", new DateTime(1926, 12, 25)),
            new Era("Heisei",  "H", new DateTime(1989,  1,  8)),
            new Era("Reiwa",   "R", new DateTime(2019,  5,  1))
        };

        /// <param name="date">Gregorian date to convert.</param>
        /// <returns>Formatted Japanese era date string.</returns>
        public static string ToJapaneseEraString(DateTime date)
        {
            // Ensure the date is not earlier than the first supported era.
            if (date < Eras[0].StartDate)
                throw new ArgumentOutOfRangeException(nameof(date), "Date is earlier than supported Japanese eras.");

            // Find the era that the date belongs to.
            Era currentEra = null;
            for (int i = Eras.Length - 1; i >= 0; i--)
            {
                if (date >= Eras[i].StartDate)
                {
                    currentEra = Eras[i];
                    break;
                }
            }

            // Compute the year within the era.
            int eraYear = date.Year - currentEra.StartDate.Year + 1;
            // Adjust for dates before the start month/day of the era year.
            if (date.Month < currentEra.StartDate.Month ||
                (date.Month == currentEra.StartDate.Month && date.Day < currentEra.StartDate.Day))
            {
                eraYear--;
            }

            // First year is represented by "元".
            string yearPart = eraYear == 1 ? "元" : eraYear.ToString();

            // Build the final string: e.g., "R3年5月1日"
            return $"{currentEra.Abbreviation}{yearPart}年{date.Month}月{date.Day}日";
        }
    }

    // Entry point of the console application.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define output file path.
                string outputPath = "JapaneseEraDate.xlsx";

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Get current date and format it using the Japanese era formatter.
                DateTime now = DateTime.Now;
                string eraString = JapaneseEraFormatter.ToJapaneseEraString(now);

                // Write the formatted date into cell A1.
                sheet.Cells["A1"].PutValue($"Current Japanese Era Date: {eraString}");

                // Save the workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook successfully saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
