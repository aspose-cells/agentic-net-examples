// Title: C# console application that prompts for a culture code and localizes an Excel workbook with Aspose.Cells
// AI Prompts: Write a C# console program that reads a culture identifier from the user, validates it with CultureInfo, assigns it to Workbook.Settings.CultureInfo, and saves the workbook using a filename that includes the culture name. | Modify an existing Aspose.Cells workbook loading routine to accept a user‑provided locale, apply the locale to the workbook's settings, and output the file with a locale‑specific suffix.
// Common Searches: how to apply a user selected culture to an Excel workbook using Aspose.Cells in C# | C# console ask for language code and save localized Excel file | set Workbook.Settings.CultureInfo based on user input Aspose.Cells | save Excel workbook with locale identifier in filename .NET | validate culture code entered by user in a C# console app
// Tags: Aspose.Cells set workbook cultureinfo | C# console apply locale to Excel workbook | localize Excel file using CultureInfo | save workbook with locale suffix | validate .NET culture code from input

using System;
using System.Globalization;
using Aspose.Cells;

namespace WorkbookLocalizationApp
{
    // The example shows a C# console app that prompts the user for a culture code, validates it, loads an Excel workbook with Aspose.Cells, sets the workbook's CultureInfo to the selected locale, and saves the file with the culture name appended to the filename.
    class Program
    {
        static void Main(string[] args)
        {
            // Prompt the user to select a target language (culture code)
            Console.WriteLine("Enter the target language code (e.g., en-US, fr-FR, de-DE):");
            string languageCode = Console.ReadLine();

            // Validate the entered culture code
            CultureInfo cultureInfo;
            try
            {
                cultureInfo = new CultureInfo(languageCode);
            }
            catch (CultureNotFoundException)
            {
                Console.WriteLine("Invalid culture code. Using invariant culture.");
                cultureInfo = CultureInfo.InvariantCulture;
            }

            // Load the workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Apply localization by setting the workbook's culture info
            workbook.Settings.CultureInfo = cultureInfo;

            // Optionally, you can also set the language for built‑in functions
            // workbook.Settings.CultureInfo = new CultureInfo(languageCode);

            // Save the localized workbook with a language‑specific file name
            string outputPath = $"output_{cultureInfo.Name}.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved with localization applied: {outputPath}");
        }
    }
}
