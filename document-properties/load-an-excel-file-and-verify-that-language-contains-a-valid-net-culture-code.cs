// Title: Validate Excel Language Property as .NET Culture Code with Aspose.Cells (C#)
// Description: Loads an Excel workbook using Aspose.Cells, reads the built‑in Language property, checks for emptiness, and attempts to instantiate a System.Globalization.CultureInfo. The program reports whether the language tag is a recognized .NET culture identifier.
// Keywords: Aspose.Cells | C# | Excel language property | built‑in document properties | CultureInfo validation | culture code check | Excel metadata | LoadOptions | Workbook language tag | .NET culture identifier
// Common Searches: Aspose.Cells read language property | C# validate Excel language tag | Check .NET culture code in Excel file | How to verify built‑in Language property Aspose | Excel document culture validation .NET
// Developer Intent: Confirm that the Language built‑in property of an Excel file contains a valid .NET culture identifier.
// Use Cases: Audit incoming Excel files for missing or empty Language tags and log warnings. | Prevent processing of workbooks with invalid culture codes in data‑import pipelines. | Integrate language‑code verification into batch jobs that enforce metadata compliance.
// AI Prompts: Generate a C# method that loads an Excel file with Aspose.Cells, reads BuiltInDocumentProperties.Language, and returns true only if the value can be parsed into a valid CultureInfo. | Provide code that throws a custom InvalidCultureException when the Language property of a workbook is not a recognized .NET culture, using Aspose.Cells and System.Globalization. | Create a reusable routine that logs each workbook’s language tag and the result of its culture‑code validation.

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsLanguageValidation
{
    // Loads an Excel workbook using Aspose.Cells, reads the built‑in Language property, checks for emptiness, and attempts to instantiate a System.Globalization.CultureInfo. The program reports whether the language tag is a recognized .NET culture identifier.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "sample.xlsx";

            // Create LoadOptions (default) and load the workbook
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Retrieve the language property from built‑in document properties
            string languageTag = workbook.BuiltInDocumentProperties.Language;

            // If the language property is empty, report it
            if (string.IsNullOrWhiteSpace(languageTag))
            {
                Console.WriteLine("The document does not contain a language tag.");
                return;
            }

            // Attempt to create a CultureInfo from the language tag
            try
            {
                CultureInfo culture = new CultureInfo(languageTag);
                Console.WriteLine($"Valid .NET culture code detected: {culture.Name}");
            }
            catch (CultureNotFoundException)
            {
                Console.WriteLine($"Invalid .NET culture code: '{languageTag}'");
            }
        }
    }
}
