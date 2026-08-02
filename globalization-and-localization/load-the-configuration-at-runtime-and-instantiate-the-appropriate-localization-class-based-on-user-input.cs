// Title: Runtime Locale Loading and Localization with Aspose.Cells (C#)
// Description: A console app that reads a locale code from the user, maps it to a full CultureInfo, configures LoadOptions, optionally loads an existing workbook, sets Workbook.Settings.LanguageCode, writes the selected locale to a cell, and saves the file as a localized Excel document.
// Keywords: Aspose.Cells CultureInfo | LoadOptions locale | Workbook Settings LanguageCode | C# Excel localization | dynamic locale mapping | runtime culture Excel | Aspose.Cells internationalization
// Common Searches: Aspose.Cells load workbook with specific CultureInfo | set workbook language code based on user locale | C# map locale to CountryCode Aspose.Cells | create new workbook when file missing Aspose.Cells | apply runtime localization to Excel with Aspose
// Developer Intent: Read a user‑provided locale at runtime, apply it to Aspose.Cells LoadOptions and WorkbookSettings, and produce a localized Excel file.
// Use Cases: Load an existing spreadsheet so dates, numbers, and currency follow the user’s culture. | Generate a fresh workbook when the source file is absent and assign the correct LanguageCode for formula behavior. | Record the applied locale in a worksheet cell for audit trails or downstream processing.
// AI Prompts: Write C# code that accepts a locale string, creates a CultureInfo, sets LoadOptions.CultureInfo, and updates Workbook.Settings.LanguageCode using Aspose.Cells. | Show how to extend the locale‑to‑CountryCode dictionary with additional languages and handle unknown locales gracefully. | Explain how Aspose.Cells formats dates, numbers, and currencies according to the CultureInfo supplied in LoadOptions.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// A console app that reads a locale code from the user, maps it to a full CultureInfo, configures LoadOptions, optionally loads an existing workbook, sets Workbook.Settings.LanguageCode, writes the selected locale to a cell, and saves the file as a localized Excel document.
public class LocalizationDemo
{
    public static void Main()
    {
        try
        {
            // Prompt user for locale code (e.g., "en", "de", "fr")
            Console.Write("Enter locale code (e.g., en, de, fr): ");
            string userLocale = Console.ReadLine()?.Trim().ToLower();

            // Validate input
            if (string.IsNullOrEmpty(userLocale))
            {
                Console.WriteLine("No locale provided. Exiting.");
                return;
            }

            // Map simple locale identifiers to full CultureInfo names
            var localeToCulture = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "ar", "ar-SA" }, { "bg", "bg-BG" }, { "ca", "ca-ES" }, { "cs", "cs-CZ" },
                { "da", "da-DK" }, { "de", "de-DE" }, { "el", "el-GR" }, { "en", "en-US" },
                { "es", "es-ES" }, { "fa", "fa-IR" }, { "fr", "fr-FR" }, { "he", "he-IL" },
                { "hi", "hi-IN" }, { "hr", "hr-HR" }, { "hu", "hu-HU" }, { "id", "id-ID" },
                { "it", "it-IT" }, { "ja", "ja-JP" }, { "ka", "ka-GE" }, { "ko", "ko-KR" },
                { "lt", "lt-LT" }, { "ms", "ms-MY" }, { "nl", "nl-NL" }, { "pl", "pl-PL" },
                { "pt", "pt-PT" }, { "ro", "ro-RO" }, { "ru", "ru-RU" }, { "sk", "sk-SK" },
                { "th", "th-TH" }, { "tr", "tr-TR" }, { "uk", "uk-UA" }, { "vi", "vi-VN" },
                { "zh", "zh-CN" }, { "zh-hant", "zh-TW" }
            };

            // Determine CultureInfo based on user input; fallback to invariant culture if unknown
            CultureInfo cultureInfo;
            if (localeToCulture.TryGetValue(userLocale, out string cultureName))
            {
                cultureInfo = new CultureInfo(cultureName);
            }
            else
            {
                Console.WriteLine($"Locale '{userLocale}' not recognized. Using invariant culture.");
                cultureInfo = CultureInfo.InvariantCulture;
            }

            // Prepare LoadOptions with the selected CultureInfo
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = cultureInfo
            };

            // Load an existing workbook if it exists; otherwise create a new one
            string inputPath = "input.xlsx";
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath, loadOptions);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Map locale to CountryCode for WorkbookSettings.LanguageCode
            var localeToCountryCode = new Dictionary<string, CountryCode>(StringComparer.OrdinalIgnoreCase)
            {
                { "en", CountryCode.USA },
                { "de", CountryCode.Germany },
                { "fr", CountryCode.France },
                { "es", CountryCode.Spain },
                { "zh", CountryCode.China },
                { "ja", CountryCode.Japan },
                { "ru", CountryCode.Russia }
                // Add more mappings as needed
            };

            if (localeToCountryCode.TryGetValue(userLocale, out CountryCode countryCode))
            {
                workbook.Settings.LanguageCode = countryCode;
            }

            // Example modification: write the selected locale into a cell
            workbook.Worksheets[0].Cells["A1"].PutValue($"Locale: {userLocale}");

            // Save the workbook with a new name
            string outputPath = "output_localized.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}' with locale '{userLocale}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
