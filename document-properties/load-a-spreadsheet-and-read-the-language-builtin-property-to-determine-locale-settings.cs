using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    public class ReadBuiltInLanguageDemo
    {
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
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Create LoadOptions instance (default constructor)
                LoadOptions loadOptions = new LoadOptions();

                // Load the workbook using the LoadOptions
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Get the built‑in document properties collection
                BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

                // Read the Language property (locale information)
                string language = builtInProps.Language;

                // Display the language value
                Console.WriteLine("Document Language (locale): " + language);
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine("An error occurred while reading the document language:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}