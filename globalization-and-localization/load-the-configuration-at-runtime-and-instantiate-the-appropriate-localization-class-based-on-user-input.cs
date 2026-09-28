// Title: Load localization mappings from a JSON file at runtime and instantiate the appropriate ILocalization implementation using reflection in C#
// AI Prompts: Write C# code that reads a JSON file containing language‑code to fully‑qualified class name mappings, asks the user for a language code, and creates the matching ILocalization object via reflection. | Update the sample to log an error and automatically fall back to EnglishLocalization when the requested language code is absent or the target type cannot be instantiated. | Enhance the program to scan a folder for additional assemblies, load them at runtime, and instantiate localization classes defined in the JSON mapping.
// Common Searches: how to map language codes to class names using a JSON file in .NET | create a localization instance from user input at runtime c# | default fallback localization when configuration entry is missing c# | load external localization assemblies dynamically based on a config file
// Tags: load localization mapping from JSON in C# | instantiate interface implementation via reflection | default language fallback strategy | dynamic assembly discovery for localization | handle missing configuration file errors

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;
using Aspose.Cells; // Aspose.Cells namespace (required by the project)

// Interface that all localization classes must implement
public interface ILocalization
{
    string GetGreeting();
}

// Example localization implementations
// Demonstrates loading a JSON configuration that maps language codes to fully‑qualified localization class names, prompting the user for a language, resolving the class with reflection, creating an ILocalization instance, and displaying its greeting. Includes error handling for missing files, JSON parsing issues, and type mismatches, and shows how to add a default fallback and extend to dynamic assembly loading.
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
    // Maps a language key (e.g., "en", "es") to the fully‑qualified class name
    public Dictionary<string, string> Mappings { get; set; }
}

public class Program
{
    // Path to the JSON configuration file (adjust as needed)
    private const string ConfigFilePath = "localizationConfig.json";

    public static void Main()
    {
        try
        {
            // 1. Load configuration at runtime
            LocalizationConfig config = LoadConfiguration(ConfigFilePath);

            // 2. Prompt user for language selection
            Console.WriteLine("Enter language code (e.g., en, es):");
            string userInput = Console.ReadLine()?.Trim().ToLower();

            // 3. Resolve the appropriate localization class name
            if (string.IsNullOrEmpty(userInput) || !config.Mappings.TryGetValue(userInput, out string className))
            {
                Console.WriteLine("Unsupported language. Falling back to English.");
                className = typeof(EnglishLocalization).FullName;
            }

            // 4. Instantiate the localization class using reflection
            ILocalization localizationInstance = CreateLocalizationInstance(className);

            // 5. Use the instantiated class
            Console.WriteLine(localizationInstance.GetGreeting());

            // Example of using Aspose.Cells after localization (optional)
            // Workbook wb = new Workbook(); // create a new workbook
            // // ... further Aspose.Cells operations ...
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        catch (JsonException ex)
        {
            Console.WriteLine($"Error parsing configuration: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Loads the JSON configuration file and deserializes it into LocalizationConfig
    private static LocalizationConfig LoadConfiguration(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException($"Configuration file not found: {path}");
        }

        string json = File.ReadAllText(path);
        var options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        };
        return JsonSerializer.Deserialize<LocalizationConfig>(json, options);
    }

    // Creates an instance of the class identified by its fully‑qualified name
    private static ILocalization CreateLocalizationInstance(string fullyQualifiedName)
    {
        // Load the current assembly (assuming the classes are in the same assembly)
        Assembly assembly = Assembly.GetExecutingAssembly();

        // Get the Type object for the class name
        Type type = assembly.GetType(fullyQualifiedName);
        if (type == null)
        {
            throw new InvalidOperationException($"Type '{fullyQualifiedName}' not found in assembly.");
        }

        // Ensure the type implements ILocalization
        if (!typeof(ILocalization).IsAssignableFrom(type))
        {
            throw new InvalidOperationException($"Type '{fullyQualifiedName}' does not implement ILocalization.");
        }

        // Create an instance using the default constructor
        return (ILocalization)Activator.CreateInstance(type);
    }
}
