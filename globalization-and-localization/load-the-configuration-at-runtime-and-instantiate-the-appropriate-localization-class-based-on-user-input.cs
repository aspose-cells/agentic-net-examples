// Title: Load a JSON localization configuration at runtime and instantiate the appropriate language class using reflection in C#
// Common Searches: how to read a JSON localization config file at runtime and create language objects with reflection in .NET | C# example for mapping language codes to fully qualified type names and instantiating them dynamically | fallback to default language mappings when localization JSON file is missing in a console app | using Type.GetType to load a class by name from a configuration and cast to an interface in C# | create ILocalization implementation based on user input language code with Activator.CreateInstance
// Tags: runtime JSON localization mapping with reflection | dynamic language class instantiation using Type.GetType | default language fallback configuration C# | ILocalization interface implementation pattern | activator createinstance for localization classes

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Reflection;

// Interface that all localization classes must implement
public interface ILocalization
{
    string GetGreeting();
}

// Example localization implementations
// The example reads a JSON file that maps language codes to fully‑qualified type names, ensures default mappings for English and French, prompts the user for a language code, and uses reflection (Type.GetType and Activator.CreateInstance) to create an ILocalization instance for the selected language, handling missing files or invalid mappings gracefully.
public class EnglishLocalization : ILocalization
{
    public string GetGreeting() => "Hello!";
}

public class FrenchLocalization : ILocalization
{
    public string GetGreeting() => "Bonjour!";
}

// Model representing the configuration file structure
public class LocalizationConfig
{
    // Maps a language code (e.g., "en", "fr") to the fully‑qualified type name of the class
    public Dictionary<string, string> LanguageMap { get; set; } = new Dictionary<string, string>();
}

public class Program
{
    // Path to the JSON configuration file (adjust as needed)
    private const string ConfigFilePath = "localizationConfig.json";

    public static void Main()
    {
        try
        {
            // Load configuration at runtime (fallback to defaults if file missing)
            LocalizationConfig config = LoadConfiguration(ConfigFilePath);

            // Prompt user for language selection
            Console.WriteLine("Enter language code (e.g., en, fr):");
            string? userInput = Console.ReadLine()?.Trim().ToLower();

            if (string.IsNullOrEmpty(userInput))
            {
                Console.WriteLine("No language code entered.");
                return;
            }

            // Instantiate the appropriate localization class based on user input
            ILocalization? localization = CreateLocalizationInstance(config, userInput);

            // Use the localization instance
            if (localization != null)
            {
                Console.WriteLine(localization.GetGreeting());
            }
            else
            {
                Console.WriteLine("Unsupported language or configuration error.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the application from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Loads the JSON configuration file and deserializes it into a LocalizationConfig object.
    // If the file does not exist or cannot be parsed, returns a config with default mappings.
    private static LocalizationConfig LoadConfiguration(string path)
    {
        var config = new LocalizationConfig();

        if (File.Exists(path))
        {
            try
            {
                string json = File.ReadAllText(path);
                var deserialized = JsonSerializer.Deserialize<LocalizationConfig>(json);
                if (deserialized?.LanguageMap != null)
                {
                    config.LanguageMap = deserialized.LanguageMap;
                }
            }
            catch
            {
                // If deserialization fails, fall back to defaults.
            }
        }

        // Ensure default mappings are present (use fully‑qualified type names)
        if (!config.LanguageMap.ContainsKey("en"))
        {
            config.LanguageMap["en"] = typeof(EnglishLocalization).FullName!;
        }
        if (!config.LanguageMap.ContainsKey("fr"))
        {
            config.LanguageMap["fr"] = typeof(FrenchLocalization).FullName!;
        }

        return config;
    }

    // Creates an instance of the localization class mapped to the supplied language code.
    private static ILocalization? CreateLocalizationInstance(LocalizationConfig config, string languageCode)
    {
        if (config.LanguageMap == null || !config.LanguageMap.TryGetValue(languageCode, out string? typeName))
        {
            return null; // No mapping found for the requested language
        }

        // Attempt to locate the type using the fully‑qualified name
        Type? targetType = Type.GetType(typeName);
        if (targetType == null)
        {
            return null; // Type not found (could be missing assembly reference)
        }

        // Ensure the type implements ILocalization and has a parameterless constructor
        if (!typeof(ILocalization).IsAssignableFrom(targetType))
        {
            return null; // Type does not implement the required interface
        }

        // Create an instance via reflection
        return Activator.CreateInstance(targetType) as ILocalization;
    }
}
