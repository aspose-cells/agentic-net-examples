// Title: Load Runtime Locale from Config and Apply Aspose.Cells Localization in C#
// Description: Read a locale string from a config.txt file, create the matching CultureInfo, configure Aspose.Cells LoadOptions, set the workbook's UI language via CountryCode, and save the localized Excel file. Includes fallback to invariant culture for unsupported locales.
// Keywords: Aspose.Cells C# | LoadOptions CultureInfo | runtime locale configuration | Excel workbook localization | CountryCode mapping | CultureInfo fallback | dynamic language settings | config file locale
// Common Searches: Aspose.Cells load workbook with cultureinfo from config | C# set workbook language code using Aspose.Cells | How to map locale string to Aspose.Cells CountryCode | LoadOptions CultureInfo example Aspose.Cells | Read locale from text file and apply to Excel workbook C#
// Developer Intent: Read a locale at runtime, convert it to .NET CultureInfo and Aspose.Cells CountryCode, apply these settings to LoadOptions and workbook UI language, and save the localized workbook.
// Use Cases: Automatically format numbers, dates, and currencies according to a user‑selected region when loading Excel files. | Display workbook UI elements (menus, messages) in the language defined by a configuration file. | Provide a safe fallback to invariant culture for unknown or misspelled locale codes, preventing runtime errors.
// AI Prompts: Generate C# code that reads a locale from a JSON configuration file and applies it to Aspose.Cells LoadOptions and workbook Settings.LanguageCode. | Extend GetCountryCodeFromLocale to include "pt-BR", "it", and "nl" with appropriate CountryCode values. | Explain best practices for validating and sanitizing locale strings before creating a CultureInfo object in Aspose.Cells.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// Read a locale string from a config.txt file, create the matching CultureInfo, configure Aspose.Cells LoadOptions, set the workbook's UI language via CountryCode, and save the localized Excel file. Includes fallback to invariant culture for unsupported locales.
static class Config
{
    // Holds the locale string for the application
    public static string Locale { get; set; } = "en";
}

class Program
{
    static void Main()
    {
        // Load locale setting from a simple configuration file at runtime
        string configPath = "config.txt";
        string locale = "en"; // Default locale

        if (File.Exists(configPath))
        {
            // Read the locale code (e.g., "en", "de", "zh-hant") and trim whitespace
            locale = File.ReadAllText(configPath).Trim();
        }

        // Assign the loaded locale to the static Config.Locale field
        Config.Locale = locale;

        // Create a CultureInfo instance based on the locale string
        CultureInfo culture = GetCultureInfoFromLocale(locale);

        // Prepare LoadOptions with the selected CultureInfo
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            CultureInfo = culture
        };

        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load an existing workbook using the configured LoadOptions
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Optionally set the workbook's UI language based on the locale
            workbook.Settings.LanguageCode = GetCountryCodeFromLocale(locale);

            // Save the workbook after applying localization settings
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Converts a locale code (e.g., "en", "zh-hant") to a .NET CultureInfo
    static CultureInfo GetCultureInfoFromLocale(string locale)
    {
        try
        {
            // Replace underscores with hyphens to match .NET culture naming
            string cultureName = locale.Replace('_', '-');
            return new CultureInfo(cultureName);
        }
        catch
        {
            // Fallback to invariant culture if the locale is not recognized
            return CultureInfo.InvariantCulture;
        }
    }

    // Maps a locale code to the corresponding Aspose.Cells CountryCode enum value
    static CountryCode GetCountryCodeFromLocale(string locale)
    {
        switch (locale)
        {
            case "en":
                return CountryCode.USA;
            case "de":
                return CountryCode.Germany;
            case "fr":
                return CountryCode.France;
            case "es":
                return CountryCode.Spain;
            case "zh":
                return CountryCode.China;
            case "zh-hant":
                return CountryCode.Taiwan;
            case "ja":
                return CountryCode.Japan;
            case "ru":
                return CountryCode.Russia;
            default:
                // Default to English (USA) if no specific mapping exists
                return CountryCode.USA;
        }
    }
}
