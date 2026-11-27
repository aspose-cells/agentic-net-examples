// Title: C# console app that prompts for a culture code and localizes an Excel workbook using Aspose.Cells
// Common Searches: C# console ask user for language code and set workbook culture using Aspose.Cells | How to apply CultureInfo to an Excel file with Aspose.Cells in .NET | Localizing Excel workbooks programmatically with Aspose.Cells and user input | Validate culture name before assigning to Aspose.Cells workbook settings | Save a localized copy of an Excel workbook after changing workbook.Settings.CultureInfo | Aspose.Cells example for dynamic language selection in a console app
// Tags: Aspose.Cells workbook cultureinfo assignment | C# console culture name validation | Excel file localization via Aspose.Cells | Programmatic language switching for Excel workbooks | Saving localized Excel workbook with Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

namespace WorkbookLocalization
{
    // The console program reads a culture identifier from the user, validates it with .NET's CultureInfo, loads an Excel file via Aspose.Cells, assigns the CultureInfo to workbook.Settings.CultureInfo, and saves the workbook as a localized copy.
    class Program
    {
        static void Main(string[] args)
        {
            // Prompt the user to select a target language (culture name)
            Console.WriteLine("Enter the target language (e.g., en-US, fr-FR, de-DE):");
            string cultureName = Console.ReadLine();

            // Validate the culture name
            CultureInfo cultureInfo;
            try
            {
                cultureInfo = new CultureInfo(cultureName);
            }
            catch (CultureNotFoundException)
            {
                Console.WriteLine("Invalid culture name. Using invariant culture.");
                cultureInfo = CultureInfo.InvariantCulture;
            }

            // Load the workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Apply localization by setting the workbook's culture info
            workbook.Settings.CultureInfo = cultureInfo;

            // Optionally, you can also set the default language for built‑in functions
            // workbook.Settings.CultureInfo = cultureInfo; // already set above

            // Save the localized workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook has been localized to {cultureInfo.Name} and saved to {outputPath}.");
        }
    }
}
