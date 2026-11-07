// Title: Convert Gregorian dates to Japanese era strings in an Excel workbook using Aspose.Cells C# API while preserving era information
// Common Searches: C# Aspose.Cells convert Excel date to Japanese era format | How to preserve Japanese era when converting dates in a workbook with Aspose.Cells | Replace Gregorian dates with Japanese calendar strings in a spreadsheet using Aspose.Cells | Using JapaneseCalendar and CultureInfo to format dates in Aspose.Cells C# example | Iterate through cells and change DateTime values to Japanese era with Aspose.Cells
// Tags: aspnet-cells japanese-era conversion | excel date localization japanese calendar | cellshelper format datetime to japanese era | c# spreadsheet era preservation aspnet-cells | japanesecalendar formatting with aspnet-cells

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The example loads an Excel workbook, scans all used cells, detects DateTime values, converts each Gregorian date to a Japanese calendar string that includes the era using JapaneseCalendar and CultureInfo, replaces the original cell content with the formatted era string, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a DateTime value
                    if (cell.Type == CellValueType.IsDateTime)
                    {
                        DateTime gregorianDate = cell.DateTimeValue;

                        // Convert Gregorian date to Japanese calendar string (preserves era)
                        string japaneseDate = ConvertToJapaneseEra(gregorianDate);

                        // Replace the original Gregorian date with the Japanese calendar string
                        cell.PutValue(japaneseDate);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to format a DateTime using the Japanese calendar (including era)
    private static string ConvertToJapaneseEra(DateTime date)
    {
        // Set up Japanese calendar with Japanese culture info
        JapaneseCalendar japaneseCalendar = new JapaneseCalendar();
        CultureInfo japaneseCulture = new CultureInfo("ja-JP");
        DateTimeFormatInfo dtfi = japaneseCulture.DateTimeFormat;
        dtfi.Calendar = japaneseCalendar;

        // Example format: "平成31年4月30日" or "令和1年5月1日"
        // Adjust the format string as needed.
        return date.ToString("gg y年M月d日", dtfi);
    }
}
