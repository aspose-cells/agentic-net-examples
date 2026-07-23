// Title: Validate Excel Workbook Language Property with AspNet Cells and .NET CultureInfo
// Description: Loads an Excel file using Aspose.Cells, reads the BuiltInDocumentProperties.Language value, and verifies that it matches a valid .NET culture identifier via CultureInfo. The sample reports missing, valid, or invalid language codes to the console.
// Keywords: Aspose.Cells language property validation | Excel BuiltInDocumentProperties Language | CultureInfo GetCultureInfo Excel | verify workbook language .NET | check Excel file culture code | C# Aspose.Cells document properties
// Common Searches: how to read language property from Excel with Aspose.Cells | validate Excel workbook language code in C# | check if Excel BuiltInDocumentProperties.Language is a valid .NET culture | Aspose.Cells verify workbook language | C# CultureInfo validation for Excel files
// Developer Intent: Confirm that the Language built‑in property of an Excel workbook contains a valid .NET culture code.
// Use Cases: Ensure localization compliance by validating language metadata before processing workbooks. | Automatically flag or reject Excel files with missing or incorrect language codes during bulk import. | Record validated language identifiers for audit trails in document management systems.
// AI Prompts: Write a C# function that returns true only when an Excel file's BuiltInDocumentProperties.Language is a valid CultureInfo. | Create unit tests for the language‑validation example covering valid, invalid, and absent language scenarios. | Generate code that scans a directory of Excel files, applies the language check, and outputs a summary of valid and invalid entries.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    // Loads an Excel file using Aspose.Cells, reads the BuiltInDocumentProperties.Language value, and verifies that it matches a valid .NET culture identifier via CultureInfo. The sample reports missing, valid, or invalid language codes to the console.
    public class VerifyWorkbookLanguage
    {
        public static void Run(string filePath)
        {
            // Ensure the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook using default LoadOptions
                LoadOptions loadOptions = new LoadOptions();
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Retrieve the language property from built‑in document properties
                string language = workbook.BuiltInDocumentProperties.Language;

                // If the language property is empty, report it
                if (string.IsNullOrWhiteSpace(language))
                {
                    Console.WriteLine("The workbook does not contain a language property.");
                    return;
                }

                // Verify the language code by creating a CultureInfo object
                try
                {
                    CultureInfo culture = CultureInfo.GetCultureInfo(language);
                    Console.WriteLine($"Valid language code found: {culture.Name}");
                }
                catch (CultureNotFoundException)
                {
                    Console.WriteLine($"Invalid language code in workbook: '{language}'");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            // Determine the workbook path: use first argument or a placeholder
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            // Execute the verification logic
            VerifyWorkbookLanguage.Run(filePath);
        }
    }
}
