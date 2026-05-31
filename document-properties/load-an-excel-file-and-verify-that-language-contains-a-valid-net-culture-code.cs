using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VerifyWorkbookLanguage
    {
        public static void Run(string filePath)
        {
            // Prevent FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook with default load options
                LoadOptions loadOptions = new LoadOptions();
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Retrieve the language property from built‑in document properties
                string language = workbook.BuiltInDocumentProperties.Language;

                // Report if the property is missing
                if (string.IsNullOrWhiteSpace(language))
                {
                    Console.WriteLine("The workbook does not contain a language property.");
                    return;
                }

                // Validate the language code
                try
                {
                    CultureInfo culture = new CultureInfo(language);
                    Console.WriteLine($"Valid language code found: {language} ({culture.DisplayName})");
                }
                catch (CultureNotFoundException)
                {
                    Console.WriteLine($"Invalid language code: {language}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }

        // Entry point for the console application
        public static void Main(string[] args)
        {
            string filePath = args.Length > 0 ? args[0] : string.Empty;

            if (string.IsNullOrWhiteSpace(filePath))
            {
                Console.Write("Enter the path to the Excel file: ");
                filePath = Console.ReadLine();
            }

            Run(filePath);
        }
    }
}