using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VerifyLanguageCultureDemo
    {
        public static void Main(string[] args)
        {
            // Determine file path from arguments or use a default placeholder
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            try
            {
                Run(filePath);
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions to prevent the program from crashing
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run(string filePath)
        {
            // Ensure the input file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Create default load options
                LoadOptions loadOptions = new LoadOptions();

                // Load the workbook with the specified load options
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Retrieve the Language property from the built‑in document properties
                string language = workbook.BuiltInDocumentProperties.Language;

                // Assume the language is valid until proven otherwise
                bool isValidCulture = true;

                // Attempt to create a CultureInfo instance; if it fails, the culture code is invalid
                try
                {
                    CultureInfo culture = new CultureInfo(language);
                }
                catch (CultureNotFoundException)
                {
                    isValidCulture = false;
                }

                // Output the result
                Console.WriteLine($"Language property: '{language}'. Valid .NET culture: {isValidCulture}");
            }
            catch (FileNotFoundException fnfEx)
            {
                // Specific handling for missing file during workbook loading
                Console.WriteLine($"Error loading workbook: {fnfEx.Message}");
            }
            catch (Exception ex)
            {
                // General exception handling for any other errors
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}