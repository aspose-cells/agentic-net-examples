// Title: Load a JSON localization config at runtime and create language-specific ILocalization instances using reflection in C#
// Common Searches: C# load localization settings from a JSON file and instantiate class by type name | How to use System.Text.Json to map language codes to localization classes in .NET | Create a dynamic ILocalization implementation based on user-entered language code | Reflection based factory for language-specific localization objects in C# | Handle missing configuration file when loading localization map in a console app
// Tags: json driven localization mapping c# | reflection localization instance creation | runtime language code to type resolution | system.text.json configuration loading c# | dynamic ILocalization factory pattern

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;

namespace LocalizationDemo
{
    // Interface that all localization classes must implement
    public interface ILocalization
    {
        string GetGreeting();
    }

    // Example localization implementations
    // The console program prompts for a language code, reads a JSON file that maps language keys to fully qualified type names, validates the input, uses reflection to instantiate the corresponding ILocalization implementation, and outputs the localized greeting, handling missing files and type errors gracefully.
    public class EnglishLocalization : ILocalization
    {
        public string GetGreeting() => "Hello!";
    }

    public class SpanishLocalization : ILocalization
    {
        public string GetGreeting() => "¡Hola!";
    }

    // Class representing the configuration file structure
    public class LocalizationConfig
    {
        // Maps a language key (e.g., "en", "es") to the fully‑qualified type name of the class
        public Dictionary<string, string> LanguageMap { get; set; } = new Dictionary<string, string>();
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Prompt user for language code
                Console.Write("Enter language code (e.g., en, es): ");
                string userInput = Console.ReadLine()?.Trim().ToLower();

                // Load configuration at runtime
                LocalizationConfig config = LoadConfiguration("localizationConfig.json");

                // Validate input
                if (string.IsNullOrEmpty(userInput) || !config.LanguageMap.ContainsKey(userInput))
                {
                    Console.WriteLine("Unsupported language. Falling back to default (English).");
                    userInput = "en";
                }

                // Instantiate the appropriate localization class
                ILocalization localization = CreateLocalizationInstance(config.LanguageMap[userInput]);

                // Use the localization instance
                Console.WriteLine(localization.GetGreeting());
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            catch (InvalidOperationException ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Loads the JSON configuration file into a LocalizationConfig object
        private static LocalizationConfig LoadConfiguration(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Configuration file not found: {filePath}");

            string json = File.ReadAllText(filePath);
            return JsonSerializer.Deserialize<LocalizationConfig>(json) 
                   ?? throw new InvalidOperationException("Failed to deserialize configuration.");
        }

        // Uses reflection to create an instance of the class whose name is provided
        private static ILocalization CreateLocalizationInstance(string typeName)
        {
            // Attempt to get the Type from the current assembly
            Type? type = Type.GetType(typeName);
            if (type == null)
                throw new InvalidOperationException($"Type '{typeName}' could not be found.");

            // Ensure the type implements ILocalization
            if (!typeof(ILocalization).IsAssignableFrom(type))
                throw new InvalidOperationException($"Type '{typeName}' does not implement ILocalization.");

            // Create an instance (assumes a parameterless constructor)
            return (ILocalization)Activator.CreateInstance(type)!;
        }
    }
}
