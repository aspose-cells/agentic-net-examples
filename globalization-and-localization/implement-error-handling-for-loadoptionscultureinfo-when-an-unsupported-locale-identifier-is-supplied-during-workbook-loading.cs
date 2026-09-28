// Title: Handle unsupported locale identifiers when assigning LoadOptions.CultureInfo during Excel workbook loading with Aspose.Cells in C#
// AI Prompts: Write a C# method that loads an Excel file with Aspose.Cells, takes a locale string, and automatically falls back to InvariantCulture if the locale is not recognized. | Show how to catch CultureNotFoundException when setting LoadOptions.CultureInfo and log a warning before opening the workbook. | Generate code that validates a locale identifier before applying it to LoadOptions.CultureInfo and supplies a default culture for invalid values.
// Common Searches: Aspose.Cells C# load workbook with CultureInfo fallback for invalid locale | How to catch CultureNotFoundException in LoadOptions when loading an Excel file | Validate locale string before assigning to LoadOptions.CultureInfo in .NET | Set default culture for unsupported locale identifiers in Aspose.Cells load options
// Tags: Aspose.Cells LoadOptions CultureInfo fallback | C# CultureNotFoundException handling workbook loading | invalid locale identifier validation Aspose.Cells | Excel workbook loading with culture settings .NET | fallback to InvariantCulture Aspose.Cells

using Aspose.Cells;
using System;
using System.Globalization;

// The example demonstrates loading an Excel workbook with Aspose.Cells while assigning a CultureInfo based on a user‑provided locale string. It catches CultureNotFoundException for unsupported locales, logs the issue, and falls back to CultureInfo.InvariantCulture before saving the workbook.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Example locale identifier (replace with actual value)
        string localeId = "invalid-locale";

        // Load workbook with culture handling
        Workbook workbook = LoadWorkbookWithCulture(inputPath, localeId);

        // Perform any workbook operations here...

        // Save the workbook (using the standard save rule)
        workbook.Save(outputPath);
    }

    /// <param name="path">Path to the workbook file.</param>
    /// <param name="localeId">Locale identifier (e.g., "en-US", "fr-FR").</param>
    /// <returns>Loaded Workbook instance.</returns>
    static Workbook LoadWorkbookWithCulture(string path, string localeId)
    {
        // Create LoadOptions (using the create rule)
        LoadOptions loadOptions = new LoadOptions();

        try
        {
            // Attempt to set the specified culture (using the rule for setting CultureInfo)
            loadOptions.CultureInfo = new CultureInfo(localeId);
        }
        catch (CultureNotFoundException ex)
        {
            // Handle unsupported locale identifier
            Console.WriteLine($"Locale '{localeId}' is not supported. Falling back to InvariantCulture.");
            Console.WriteLine($"Error details: {ex.Message}");

            // Fallback to a safe default culture
            loadOptions.CultureInfo = CultureInfo.InvariantCulture;
        }

        // Load the workbook with the configured LoadOptions (using the load rule)
        Workbook workbook = new Workbook(path, loadOptions);
        return workbook;
    }
}
