// Title: Convert Gregorian dates to Japanese era strings in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, scans every worksheet, finds cells of type DateTime, converts each Gregorian date to a Japanese era representation (e.g., "令和3年5月10日"), and writes the formatted string back to the cell. | Show how to apply the custom Japanese locale number format "[$-ja-JP]ggge\"年\"m\"月\"d\"日\";@" to cells after converting dates with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# convert Excel Gregorian dates to Japanese era format | how to display Japanese era dates in an Excel workbook using Aspose.Cells | C# replace DateTime cells with Japanese calendar strings in .xlsx | apply Japanese custom number format to Excel cells with Aspose.Cells .NET | iterate all worksheets and convert dates to Japanese calendar using Aspose.Cells
// Tags: convert Gregorian dates to Japanese era Aspose.Cells | apply Japanese custom number format Excel .NET | iterate workbook cells date conversion C# | JapaneseCalendar API usage with Aspose.Cells | preserve era information in Excel cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel file with Aspose.Cells, iterates through each worksheet's used range, detects cells containing DateTime values, converts each Gregorian date to a Japanese era string (e.g., "令和3年5月10日"), replaces the cell value with the formatted string, applies a Japanese locale custom number format to keep the display consistent, and saves the updated workbook.
class JapaneseCalendarConverter
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Get the used range of the worksheet
                AsposeRange usedRange = cells.MaxDisplayRange;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Loop through each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only cells that contain a DateTime value
                        if (cell.Type == CellValueType.IsDateTime)
                        {
                            DateTime gregorianDate = cell.DateTimeValue;

                            // Convert Gregorian date to Japanese era information
                            JapaneseCalendar jpCal = new JapaneseCalendar();
                            int era = jpCal.GetEra(gregorianDate);
                            string eraName = new CultureInfo("ja-JP").DateTimeFormat.GetEraName(era);
                            int yearOfEra = jpCal.GetYear(gregorianDate);

                            // Build the Japanese calendar string (e.g., "令和3年5月10日")
                            string japaneseDate = $"{eraName}{yearOfEra}年{gregorianDate.Month}月{gregorianDate.Day}日";

                            // Replace the cell value with the Japanese calendar string
                            cell.PutValue(japaneseDate);

                            // Apply a custom number format to keep the display consistent
                            Style style = cell.GetStyle();
                            style.Custom = "[$-ja-JP]ggge\"年\"m\"月\"d\"日\";@";
                            cell.SetStyle(style);
                        }
                    }
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
