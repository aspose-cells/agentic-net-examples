// Title: Automatically convert Gregorian dates to Japanese era format across all cells in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# iterate through workbook cells and replace dates with Japanese era | How to map Gregorian dates to Reiwa, Heisei, Showa eras in an Excel file using .NET | Bulk update Excel date cells to Japanese era strings with Aspose.Cells library | C# code to detect DateTime cells and format them as Japanese era in .xlsx | Convert Excel dates to Japanese calendar automatically based on year using Aspose.Cells
// Tags: Aspose.Cells bulk date conversion to Japanese era | Excel workbook Japanese calendar formatting C# | iterate cells replace DateTime with era string Aspose.Cells | Japanese era selection by year .NET | range-based date processing in Excel using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace JapaneseEraDateConverter
{
    // Alias to avoid ambiguity with System.Range
    using AsposeRange = Aspose.Cells.Range;

    // The program loads an Excel workbook, iterates through every worksheet and each cell in the used range, detects cells containing DateTime values, determines the appropriate Japanese era (Meiji, Taisho, Showa, Heisei, Reiwa) based on the date, formats the date as "Era Year-Month-Day" (using "Gannen" for the first year), replaces the original value with this string, and saves the modified workbook to a new file.
    class Program
    {
        // Mapping of Japanese eras with their start dates (inclusive)
        private static readonly List<JapaneseEra> Eras = new List<JapaneseEra>
        {
            new JapaneseEra("Reiwa", new DateTime(2019, 5, 1)),
            new JapaneseEra("Heisei", new DateTime(1989, 1, 8)),
            new JapaneseEra("Showa", new DateTime(1926, 12, 25)),
            new JapaneseEra("Taisho", new DateTime(1912, 7, 30)),
            new JapaneseEra("Meiji", new DateTime(1868, 9, 8))
        };

        static void Main(string[] args)
        {
            const string inputPath = "InputDates.xlsx";
            const string outputPath = "OutputDates_JapaneseEra.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    // MaxDisplayRange returns a Range object (aliased as AsposeRange)
                    AsposeRange usedRange = cells.MaxDisplayRange;

                    // Determine the bounds of the used range
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
                                DateTime originalDate = cell.DateTimeValue;
                                string eraString = ConvertToJapaneseEra(originalDate);
                                // Replace the cell value with the formatted era string
                                cell.PutValue(eraString);
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Converts a Gregorian date to a Japanese era string (e.g., "Reiwa 3-04-15")
        private static string ConvertToJapaneseEra(DateTime date)
        {
            foreach (var era in Eras)
            {
                if (date >= era.StartDate)
                {
                    int eraYear = date.Year - era.StartDate.Year + 1;
                    // For the first year of an era, Japanese often use "Gannen" instead of "1"
                    string yearPart = eraYear == 1 ? "Gannen" : eraYear.ToString();
                    return $"{era.Name} {yearPart}-{date.Month:D2}-{date.Day:D2}";
                }
            }

            // If the date is before Meiji era, return the original Gregorian representation
            return date.ToString("yyyy-MM-dd");
        }

        // Simple struct to hold era information
        private struct JapaneseEra
        {
            public string Name { get; }
            public DateTime StartDate { get; }

            public JapaneseEra(string name, DateTime startDate)
            {
                Name = name;
                StartDate = startDate;
            }
        }
    }
}
