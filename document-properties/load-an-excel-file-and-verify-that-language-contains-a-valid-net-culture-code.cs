using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLanguageValidation
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            // Create LoadOptions (auto-detect format)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);

            // Load the workbook with the specified LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Retrieve the language property from built‑in document properties
            string languageCode = workbook.BuiltInDocumentProperties.Language;

            // Verify that the language string represents a valid .NET culture
            try
            {
                CultureInfo culture = new CultureInfo(languageCode);
                Console.WriteLine($"Valid culture code detected: {culture.Name}");
            }
            catch (CultureNotFoundException)
            {
                Console.WriteLine($"Invalid or missing culture code: '{languageCode}'");
            }
        }
    }
}