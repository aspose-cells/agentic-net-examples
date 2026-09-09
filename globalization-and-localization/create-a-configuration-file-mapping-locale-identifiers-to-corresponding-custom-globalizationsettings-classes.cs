// Title: Generate a JSON locale configuration file in C# by mapping culture identifiers to custom GlobalizationSettings objects
// AI Prompts: Create a C# program that builds a Dictionary<string, GlobalizationSettings> for multiple cultures and writes the content to an indented JSON file using System.Text.Json. | Add a new locale entry (e.g., es-ES) with its date format, decimal separator, and currency symbol, then regenerate the LocaleSettings.json file. | Configure JsonSerializerOptions to produce camel‑case property names while preserving dictionary keys, and serialize the locale mapping to a formatted JSON file.
// Common Searches: how to serialize a dictionary of custom objects to a formatted JSON file in C# | C# create locale settings JSON file with culture codes and custom date/number formats | example of writing a globalization configuration file per culture using System.Text.Json | store date format decimal separator and currency symbol per locale in a .NET JSON config
// Tags: dictionary to JSON serialization System.Text.Json C# | locale configuration file generation .NET | custom GlobalizationSettings class JSON output | write indented JSON file to application base directory | add culture-specific formatting settings C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace LocaleConfiguration
{
    // Custom settings for globalization per locale
    // The example defines a GlobalizationSettings class containing DateFormat, DecimalSeparator, and CurrencySymbol properties. It creates a Dictionary<string, GlobalizationSettings> that maps locale identifiers (e.g., "en-US", "fr-FR") to corresponding settings, serializes the dictionary to a pretty‑printed JSON string with System.Text.Json, and writes the result to a LocaleSettings.json file placed in the application's base directory.
    public class GlobalizationSettings
    {
        // Example: date format pattern
        public string DateFormat { get; set; }

        // Example: decimal separator
        public string DecimalSeparator { get; set; }

        // Example: currency symbol
        public string CurrencySymbol { get; set; }

        // Additional custom settings can be added here
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Build the mapping of locale identifiers to their settings
            var localeSettings = new Dictionary<string, GlobalizationSettings>
            {
                // English - United States
                ["en-US"] = new GlobalizationSettings
                {
                    DateFormat = "MM/dd/yyyy",
                    DecimalSeparator = ".",
                    CurrencySymbol = "$"
                },

                // French - France
                ["fr-FR"] = new GlobalizationSettings
                {
                    DateFormat = "dd/MM/yyyy",
                    DecimalSeparator = ",",
                    CurrencySymbol = "€"
                },

                // German - Germany
                ["de-DE"] = new GlobalizationSettings
                {
                    DateFormat = "dd.MM.yyyy",
                    DecimalSeparator = ",",
                    CurrencySymbol = "€"
                },

                // Japanese - Japan
                ["ja-JP"] = new GlobalizationSettings
                {
                    DateFormat = "yyyy/MM/dd",
                    DecimalSeparator = ".",
                    CurrencySymbol = "¥"
                }
            };

            // Serialize the dictionary to JSON with indented formatting
            var jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true,
                // Preserve the dictionary keys as they are
                Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
            };

            string json = JsonSerializer.Serialize(localeSettings, jsonOptions);

            // Define the output configuration file path
            string configFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LocaleSettings.json");

            // Write the JSON content to the file
            File.WriteAllText(configFilePath, json);

            Console.WriteLine($"Locale configuration file created at: {configFilePath}");
        }
    }
}
