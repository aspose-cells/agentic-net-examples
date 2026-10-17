// Title: Read the "Language" named range from an Excel workbook and validate it as a .NET culture code using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read named range value and verify CultureInfo | How to check if an Excel cell contains a valid .NET culture identifier with Aspose.Cells | C# load workbook and validate language code from a named range using Aspose.Cells | Fallback to cell A1 when a named range is missing in Aspose.Cells example | Validate culture code from Excel file using CultureInfo.GetCultureInfo in C#
// Tags: range retrieval Aspose.Cells | culture identifier validation .NET | workbook loading Aspose.Cells C# | fallback cell A1 when range absent | CultureInfo.GetCultureInfo usage C#

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, attempts to read the value of a named range called "Language" (using cell A1 as a fallback), and then uses CultureInfo.GetCultureInfo to determine whether the retrieved string is a valid .NET culture identifier, outputting the result or an error message.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File '{inputPath}' not found.");
                return;
            }

            // Load the Excel workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            var worksheet = workbook.Worksheets[0];

            // Try to obtain the value from a named range called "Language"
            string languageCode = string.Empty;
            var namedRange = workbook.Worksheets.GetRangeByName("Language");
            if (namedRange != null && namedRange.Value != null)
            {
                languageCode = namedRange.Value.ToString();
            }
            else
            {
                // Fallback: read the value from cell A1 (or any specific cell you expect)
                languageCode = worksheet.Cells["A1"].StringValue;
            }

            // Verify that the retrieved string is a valid .NET culture code
            try
            {
                CultureInfo culture = CultureInfo.GetCultureInfo(languageCode);
                Console.WriteLine($"Valid culture code: {culture.Name}");
            }
            catch (CultureNotFoundException)
            {
                Console.WriteLine($"Invalid culture code: '{languageCode}'");
            }
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected runtime errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
