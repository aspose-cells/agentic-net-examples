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
            string inputPath = "input.xlsx";

            // Create LoadOptions (using the provided LoadOptions class)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);

            // Load the workbook with the specified LoadOptions
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Retrieve the document's language property
            string language = workbook.BuiltInDocumentProperties.Language;

            // Verify that the language string represents a valid .NET culture code
            bool isValidCulture = false;
            try
            {
                // Attempt to create a CultureInfo instance; will throw if invalid
                CultureInfo ci = new CultureInfo(language);
                isValidCulture = true;
                Console.WriteLine($"Language '{language}' is a valid .NET culture: {ci.DisplayName}");
            }
            catch (CultureNotFoundException)
            {
                Console.WriteLine($"Language '{language}' is NOT a valid .NET culture code.");
            }

            // (Optional) Save the workbook to demonstrate usage of the save rule
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}