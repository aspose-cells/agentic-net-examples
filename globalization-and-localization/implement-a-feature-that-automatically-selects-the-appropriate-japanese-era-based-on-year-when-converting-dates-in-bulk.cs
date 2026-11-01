// Title: Bulk convert Excel DateTime cells to Japanese era strings using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells bulk replace DateTime cells with Japanese era format | How to automatically select Japanese era based on year when converting Excel dates | Iterate through Excel worksheet cells and format dates as Meiji, Taisho, Showa, Heisei, Reiwa using Aspose.Cells | Convert Excel dates to Japanese calendar strings in a .NET application | Aspose.Cells example for Japanese era localization of Excel dates
// Tags: bulk japanese era conversion aspocells | excel datetime to japanese calendar c# | aspocells date formatting japanese era | iterate worksheet cells replace datetime with era | japanese era localization aspocells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, scans every used cell, detects DateTime values, determines the correct Japanese era (Meiji, Taisho, Showa, Heisei, Reiwa) based on the year, replaces the original date with a formatted era string, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Process the first worksheet (adjust as needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through all used cells
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell contains a DateTime value
                    if (cell.Type == CellValueType.IsDateTime)
                    {
                        DateTime dt = cell.DateTimeValue; // Correct property for Aspose.Cells

                        // Convert the date to the appropriate Japanese era string
                        string eraDate = ConvertToJapaneseEra(dt);

                        // Replace the cell value with the era formatted string
                        cell.PutValue(eraDate);
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Returns a formatted string with the correct Japanese era based on the year
    static string ConvertToJapaneseEra(DateTime dt)
    {
        // Era boundaries (inclusive start year)
        // Meiji:   1868‑1912
        // Taisho:  1912‑1926
        // Showa:   1926‑1989
        // Heisei:  1989‑2019
        // Reiwa:   2019‑present

        if (dt.Year >= 2019) // Reiwa
        {
            int eraYear = dt.Year - 2018; // Reiwa 1 started in 2019
            return $"Reiwa {eraYear}年{dt.Month}月{dt.Day}日";
        }
        else if (dt.Year >= 1989) // Heisei
        {
            int eraYear = dt.Year - 1988;
            return $"Heisei {eraYear}年{dt.Month}月{dt.Day}日";
        }
        else if (dt.Year >= 1926) // Showa
        {
            int eraYear = dt.Year - 1925;
            return $"Showa {eraYear}年{dt.Month}月{dt.Day}日";
        }
        else if (dt.Year >= 1912) // Taisho
        {
            int eraYear = dt.Year - 1911;
            return $"Taisho {eraYear}年{dt.Month}月{dt.Day}日";
        }
        else if (dt.Year >= 1868) // Meiji
        {
            int eraYear = dt.Year - 1867;
            return $"Meiji {eraYear}年{dt.Month}月{dt.Day}日";
        }
        else
        {
            // Dates before Meiji are returned in a standard format
            return dt.ToString("yyyy/MM/dd");
        }
    }
}
