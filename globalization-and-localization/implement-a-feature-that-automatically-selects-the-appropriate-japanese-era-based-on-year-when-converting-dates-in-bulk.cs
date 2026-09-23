// Title: Automatically replace Gregorian dates with Japanese era strings in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, scans every cell for DateTime values, converts each date to the appropriate Japanese era (Meiji, Taisho, Showa, Heisei, Reiwa) using a lookup of era start dates, writes the era string back as text, and saves the workbook. | Create a reusable C# method that receives a DateTime and returns a formatted Japanese era string (e.g., "R3/04/01"), then apply this method to all date cells in a worksheet loaded with Aspose.Cells while preserving text formatting.
// Common Searches: Aspose.Cells C# bulk convert Excel dates to Japanese era format | How to map Gregorian dates to Japanese era in a spreadsheet using .NET | Replace date cells with era strings in an Excel file programmatically | C# iterate over used range in Aspose.Cells and change date format to era | Japanese era conversion for Excel dates with Aspose.Cells library
// Tags: bulk date conversion to Japanese era using Aspose.Cells | Japanese era formatting for Excel cells in .NET | replace DateTime cells with era strings Aspose.Cells | era lookup based on Gregorian year C# | text style assignment for converted dates Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx workbook with Aspose.Cells, iterates through the used range, detects DateTime cells, converts each date to the correct Japanese era string (e.g., "R3/04/01") based on predefined era start dates, writes the result back as text with a text number format, and saves the modified workbook to output.xlsx.
class JapaneseEraDateConverter
{
    // Mapping of Japanese eras with their start dates
    private static readonly (string Era, DateTime Start)[] Eras = new (string, DateTime)[]
    {
        ("M", new DateTime(1868, 9, 8)),   // Meiji
        ("T", new DateTime(1912, 7, 30)),  // Taisho
        ("S", new DateTime(1926, 12, 25)), // Showa
        ("H", new DateTime(1989, 1, 8)),   // Heisei
        ("R", new DateTime(2019, 5, 1))    // Reiwa
    };

    // Convert a Gregorian date to Japanese era string (e.g., "R3/04/01")
    private static string ConvertToJapaneseEra(DateTime date)
    {
        // Find the era that the date belongs to (latest start date <= date)
        for (int i = Eras.Length - 1; i >= 0; i--)
        {
            if (date >= Eras[i].Start)
            {
                int eraYear = date.Year - Eras[i].Start.Year + 1;
                // Japanese era year 1 is often written as "1" (not "0")
                return $"{Eras[i].Era}{eraYear}/{date.Month:D2}/{date.Day:D2}";
            }
        }

        // If date is before Meiji era, return the original Gregorian representation
        return date.ToString("yyyy/MM/dd");
    }

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

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(inputPath);

            // Assume processing the first worksheet; adjust as needed
            Worksheet sheet = workbook.Worksheets[0];

            // Get the used range of cells
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            // Iterate through each cell in the used range
            foreach (Cell cell in usedRange)
            {
                // Check if the cell contains a date value
                if (cell.Type == CellValueType.IsDateTime)
                {
                    DateTime originalDate = cell.DateTimeValue;
                    // Convert to Japanese era format
                    string eraString = ConvertToJapaneseEra(originalDate);
                    // Replace cell value with the era string
                    cell.PutValue(eraString);
                    // Set the style to Text to preserve formatting
                    Style style = cell.GetStyle();
                    style.Number = 49; // Text format
                    cell.SetStyle(style);
                }
            }

            // Save the modified workbook (save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
