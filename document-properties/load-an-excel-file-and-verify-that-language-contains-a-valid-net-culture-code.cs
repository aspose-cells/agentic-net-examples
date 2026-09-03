// Title: Validate a .NET culture code stored in the 'Language' named range (or A1) of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# function with Aspose.Cells that loads an Excel workbook, extracts the value of the 'Language' range (or cell A1), and returns a boolean indicating whether the value matches a .NET CultureInfo. | Refactor the example to throw a custom exception when the culture code is invalid instead of writing messages to the console. | Add detailed logging to capture cases where the named range is missing and when an invalid culture identifier is detected.
// Common Searches: retrieve language code from Excel using Aspose.Cells and check if it is a valid .NET culture | how to extract a language code from an Excel file with Aspose.Cells and test it against CultureInfo | fallback to cell A1 when named range not found Aspose.Cells C# example | validate Excel workbook language setting against CultureInfo in C#
// Tags: Aspose.Cells read range by name | Aspose.Cells validate .NET CultureInfo | Excel language cell default handling | C# culture code verification from workbook | Workbook property validation using Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, attempts to read the value of a named range called "Language" (using cell A1 as a fallback), and determines whether the extracted string is a valid .NET culture identifier, outputting the result or error messages to the console.
class LanguageCultureValidator
{
    // Loads an Excel file, reads the cell named "Language" and checks if it contains a valid .NET culture code.
    public static void ValidateLanguageCulture(string excelFilePath)
    {
        if (!File.Exists(excelFilePath))
        {
            Console.WriteLine($"File not found: {excelFilePath}");
            return;
        }

        try
        {
            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(excelFilePath);

            // Try to get the named range "Language"
            Cell languageCell = null;
            try
            {
                // GetRangeByName returns an Aspose.Cells.Range
                Aspose.Cells.Range languageRange = workbook.Worksheets.GetRangeByName("Language");
                if (languageRange != null && languageRange.RowCount > 0 && languageRange.ColumnCount > 0)
                {
                    languageCell = languageRange[0, 0];
                }
            }
            catch (Exception)
            {
                // Named range not found or other error; ignore and fallback
            }

            // Fallback to cell A1 of the first worksheet if needed
            if (languageCell == null)
            {
                languageCell = workbook.Worksheets[0].Cells["A1"];
            }

            // Retrieve the cell value as a trimmed string
            string languageValue = languageCell?.StringValue?.Trim();

            if (string.IsNullOrEmpty(languageValue))
            {
                Console.WriteLine("The 'Language' cell is empty.");
                return;
            }

            // Verify that the value is a valid .NET culture code
            try
            {
                CultureInfo culture = CultureInfo.GetCultureInfo(languageValue);
                Console.WriteLine($"Valid culture code: {culture.Name} - {culture.DisplayName}");
            }
            catch (CultureNotFoundException)
            {
                Console.WriteLine($"Invalid culture code: '{languageValue}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the file: {ex.Message}");
        }
    }

    // Example usage
    static void Main(string[] args)
    {
        try
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the Excel file as an argument.");
                return;
            }

            string filePath = args[0];
            ValidateLanguageCulture(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
