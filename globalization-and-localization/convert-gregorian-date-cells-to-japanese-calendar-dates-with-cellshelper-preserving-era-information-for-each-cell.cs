// Title: Convert Gregorian date cells to Japanese era strings in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert Excel date cells to Japanese era format | replace Gregorian dates with Japanese calendar strings in a .NET workbook | format Excel cells using JapaneseCalendar with Aspose.Cells | C# read Excel dates and output era name using Aspose.Cells | save Excel file with Japanese era dates using Aspose.Cells .NET
// Tags: convert Gregorian dates to Japanese era Aspose.Cells | JapaneseCalendar date formatting Excel .NET | replace Excel cell date with era string C# | cultureinfo Japanese calendar Aspose.Cells | localize Excel dates to Japanese calendar .NET

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, scans all used cells, identifies those containing Gregorian dates, converts each date to a Japanese era string (e.g., "令和3年5月1日") using a CultureInfo configured with JapaneseCalendar, writes the formatted string back to the cell, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // CultureInfo for Japanese calendar formatting
            CultureInfo japaneseCulture = new CultureInfo("ja-JP");
            japaneseCulture.DateTimeFormat.Calendar = new JapaneseCalendar();

            // Iterate through all used cells in the worksheet
            foreach (Cell cell in sheet.Cells)
            {
                // Process only cells that contain a Gregorian date
                if (cell.Type == CellValueType.IsDateTime)
                {
                    DateTime gregorianDate = cell.DateTimeValue;

                    // Convert to Japanese era string, e.g., "令和3年5月1日"
                    string japaneseDate = gregorianDate.ToString("ggyy年M月d日", japaneseCulture);

                    // Replace the cell value with the Japanese calendar string
                    cell.PutValue(japaneseDate);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
