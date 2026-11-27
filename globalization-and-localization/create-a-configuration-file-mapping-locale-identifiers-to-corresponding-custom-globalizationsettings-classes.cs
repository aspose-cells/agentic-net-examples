// Title: Generate a JSON file that maps locale identifiers to custom GlobalizationSettings objects using System.Text.Json in C#
// Common Searches: how to serialize a Dictionary<string, GlobalizationSettings> to JSON in .NET | C# write locale configuration to a JSON file with System.Text.Json | create culture-specific settings file programmatically in .NET Core | store date, time, number formats per culture in a JSON config using C# | generate globalization configuration file at runtime in a .NET application
// Tags: serialize dictionary of custom objects to JSON with System.Text.Json | culture identifier to globalization settings mapping in C# | write JSON configuration file to application base directory | custom GlobalizationSettings class for locale formatting | pretty‑printed JSON output for localization config

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace LocalizationConfig
{
    // Custom settings for globalization per locale
    // The example defines a GlobalizationSettings class, builds a Dictionary<string, GlobalizationSettings> containing locale entries such as en-US, fr-FR, and de-DE, and uses System.Text.Json with indented formatting to serialize the dictionary to a JSON file named globalizationConfig.json placed in the application's base directory.
    public class GlobalizationSettings
    {
        public string DateFormat { get; set; }
        public string TimeFormat { get; set; }
        public string NumberDecimalSeparator { get; set; }
        public string NumberGroupSeparator { get; set; }
        public string CurrencySymbol { get; set; }
        // Add other locale‑specific properties as needed
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Build the mapping of locale identifiers to their settings
            var localeSettings = new Dictionary<string, GlobalizationSettings>
            {
                {
                    "en-US", new GlobalizationSettings
                    {
                        DateFormat = "MM/dd/yyyy",
                        TimeFormat = "hh:mm tt",
                        NumberDecimalSeparator = ".",
                        NumberGroupSeparator = ",",
                        CurrencySymbol = "$"
                    }
                },
                {
                    "fr-FR", new GlobalizationSettings
                    {
                        DateFormat = "dd/MM/yyyy",
                        TimeFormat = "HH:mm",
                        NumberDecimalSeparator = ",",
                        NumberGroupSeparator = " ",
                        CurrencySymbol = "€"
                    }
                },
                {
                    "de-DE", new GlobalizationSettings
                    {
                        DateFormat = "dd.MM.yyyy",
                        TimeFormat = "HH:mm",
                        NumberDecimalSeparator = ",",
                        NumberGroupSeparator = ".",
                        CurrencySymbol = "€"
                    }
                }
                // Add additional locales here
            };

            // Serialize the dictionary to a JSON configuration file
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                // Preserve the dictionary keys as they are
                Converters = { new JsonStringEnumConverter() }
            };

            string json = JsonSerializer.Serialize(localeSettings, options);
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "globalizationConfig.json");
            File.WriteAllText(configPath, json);

            Console.WriteLine($"Configuration file written to: {configPath}");
        }
    }
}
