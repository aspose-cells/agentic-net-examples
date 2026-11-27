// Title: Subclass Aspose.Cells Workbook in C# to get culture-specific true/false strings
// Common Searches: C# Aspose.Cells how to get localized boolean text in Excel workbook | override GetBooleanString method for French German Spanish in Aspose.Cells | create a workbook subclass that returns true/false strings based on CultureInfo | Aspose.Cells write localized boolean values to cells using custom workbook class | example of culture-aware boolean string conversion with Aspose.Cells C#
// Tags: custom workbook subclass localization Aspose.Cells | culture-specific boolean string conversion C# | Aspose.Cells GetBooleanString override | localized true false Excel output | CultureInfo based Excel boolean representation

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// Subclass Workbook to provide locale‑specific boolean strings
// Demonstrates how to inherit from Aspose.Cells.Workbook in C# to create a LocalizedWorkbook that stores a CultureInfo and overrides a GetBooleanString method. The method returns true/false strings in French, German, Spanish, or default English. The example writes boolean values and their localized text to cells and saves the workbook as an Excel file.
public class LocalizedWorkbook : Workbook
{
    private readonly CultureInfo _culture;

    public LocalizedWorkbook(CultureInfo culture) : base()
    {
        _culture = culture;
    }

    // Returns a localized representation of a Boolean value
    public string GetBooleanString(bool value)
    {
        // Simple example localization; replace with resource lookup as needed
        if (_culture.Name.Equals("fr-FR", StringComparison.OrdinalIgnoreCase))
            return value ? "VRAI" : "FAUX";
        if (_culture.Name.Equals("de-DE", StringComparison.OrdinalIgnoreCase))
            return value ? "WAHR" : "FALSCH";
        if (_culture.Name.Equals("es-ES", StringComparison.OrdinalIgnoreCase))
            return value ? "VERDADERO" : "FALSO";

        // Default English
        return value ? "TRUE" : "FALSE";
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Select the desired locale
            CultureInfo locale = new CultureInfo("fr-FR");

            // Create workbook using the localized subclass
            LocalizedWorkbook workbook = new LocalizedWorkbook(locale);

            // Add data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(true);
            sheet.Cells["A2"].PutValue(false);

            // Use the localized method to get Boolean strings
            string trueStr = workbook.GetBooleanString(true);   // "VRAI"
            string falseStr = workbook.GetBooleanString(false); // "FAUX"

            // Write the localized strings to the sheet for verification
            sheet.Cells["B1"].PutValue(trueStr);
            sheet.Cells["B2"].PutValue(falseStr);

            // Determine output file path
            string outputPath = "LocalizedBoolean.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
