using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    public class ReadBuiltInLanguageProperty
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the Excel file to be loaded
            string filePath = "sample.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Create LoadOptions using the default constructor
                LoadOptions loadOptions = new LoadOptions();

                // Load the workbook with the specified LoadOptions
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Access the built‑in document properties collection
                BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

                // Read the Language property which indicates the locale settings of the file
                string language = builtInProps.Language;

                // Output the language value
                Console.WriteLine($"Built‑in Language property: {language}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}