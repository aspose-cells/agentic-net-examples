// Title: Generate and use a JSON file that maps locale codes to custom GlobalizationSettings objects in C#
// Common Searches: C# serialize a dictionary of custom objects to JSON with System.Text.Json preserving keys | How to store culture-specific date and number formats in a JSON configuration file in .NET | Load locale-specific globalization settings from a JSON file using C# | Create a case-insensitive dictionary of locale identifiers to formatting settings in .NET | Read and write custom GlobalizationSettings per culture with System.Text.Json | Save and retrieve currency format settings for multiple cultures in a .NET application
// Tags: dictionary serialization to JSON with System.Text.Json | culture to formatting settings mapping .NET | custom GlobalizationSettings class C# | locale identifier JSON configuration file | case-insensitive dictionary for culture codes | currency and date format configuration .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace LocalizationConfig
{
    // Custom settings for globalization per locale
    // The example defines a GlobalizationSettings class and builds a case‑insensitive Dictionary that maps locale identifiers (e.g., "en-US", "fr-FR") to formatting settings. It demonstrates saving this dictionary to an indented JSON file using System.Text.Json with camelCase property names, and loading the file back into a dictionary while handling missing files and case‑insensitive property names.
    public class GlobalizationSettings
    {
        public string DateFormat { get; set; }
        public string NumberDecimalSeparator { get; set; }
        public string NumberGroupSeparator { get; set; }
        public string CurrencySymbol { get; set; }
        public string CurrencyFormat { get; set; }
    }

    class Program
    {
        // Path to the configuration file
        private const string ConfigFilePath = "globalizationConfig.json";

        static void Main(string[] args)
        {
            // Create sample configuration
            var config = CreateSampleConfiguration();

            // Save configuration to file
            SaveConfiguration(config, ConfigFilePath);
            Console.WriteLine($"Configuration saved to {ConfigFilePath}");

            // Load configuration from file
            var loadedConfig = LoadConfiguration(ConfigFilePath);
            Console.WriteLine("Loaded configuration:");
            foreach (var kvp in loadedConfig)
            {
                Console.WriteLine($"Locale: {kvp.Key}");
                Console.WriteLine($"  DateFormat: {kvp.Value.DateFormat}");
                Console.WriteLine($"  DecimalSeparator: {kvp.Value.NumberDecimalSeparator}");
                Console.WriteLine($"  GroupSeparator: {kvp.Value.NumberGroupSeparator}");
                Console.WriteLine($"  CurrencySymbol: {kvp.Value.CurrencySymbol}");
                Console.WriteLine($"  CurrencyFormat: {kvp.Value.CurrencyFormat}");
            }
        }

        // Creates a dictionary mapping locale identifiers to their settings
        private static Dictionary<string, GlobalizationSettings> CreateSampleConfiguration()
        {
            var dict = new Dictionary<string, GlobalizationSettings>(StringComparer.OrdinalIgnoreCase)
            {
                {
                    "en-US", new GlobalizationSettings
                    {
                        DateFormat = "MM/dd/yyyy",
                        NumberDecimalSeparator = ".",
                        NumberGroupSeparator = ",",
                        CurrencySymbol = "$",
                        CurrencyFormat = "¤#,##0.00"
                    }
                },
                {
                    "en-GB", new GlobalizationSettings
                    {
                        DateFormat = "dd/MM/yyyy",
                        NumberDecimalSeparator = ".",
                        NumberGroupSeparator = ",",
                        CurrencySymbol = "£",
                        CurrencyFormat = "¤#,##0.00"
                    }
                },
                {
                    "fr-FR", new GlobalizationSettings
                    {
                        DateFormat = "dd/MM/yyyy",
                        NumberDecimalSeparator = ",",
                        NumberGroupSeparator = " ",
                        CurrencySymbol = "€",
                        CurrencyFormat = "#,##0.00 ¤"
                    }
                },
                {
                    "de-DE", new GlobalizationSettings
                    {
                        DateFormat = "dd.MM.yyyy",
                        NumberDecimalSeparator = ",",
                        NumberGroupSeparator = ".",
                        CurrencySymbol = "€",
                        CurrencyFormat = "#,##0.00 ¤"
                    }
                }
            };

            return dict;
        }

        // Serializes the configuration dictionary to a JSON file
        private static void SaveConfiguration(Dictionary<string, GlobalizationSettings> config, string filePath)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                // Preserve dictionary keys as they are
                DictionaryKeyPolicy = null,
                // Ensure property names are camelCase in JSON (optional)
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            string json = JsonSerializer.Serialize(config, options);
            File.WriteAllText(filePath, json);
        }

        // Deserializes the configuration dictionary from a JSON file
        private static Dictionary<string, GlobalizationSettings> LoadConfiguration(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Configuration file not found: {filePath}");

            string json = File.ReadAllText(filePath);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };

            var config = JsonSerializer.Deserialize<Dictionary<string, GlobalizationSettings>>(json, options);
            return config ?? new Dictionary<string, GlobalizationSettings>();
        }
    }
}
