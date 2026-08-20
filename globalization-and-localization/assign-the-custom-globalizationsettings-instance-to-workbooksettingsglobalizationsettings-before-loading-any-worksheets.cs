// Title: Assign a Custom GlobalizationSettings Object to Aspose.Cells Workbook in C# Before Accessing Worksheets
// Description: This example shows how to create a subclass of Aspose.Cells.GlobalizationSettings, override GetBooleanValueString and GetErrorValueString, and attach the custom instance to Workbook.Settings.GlobalizationSettings before any worksheet is read or written. The code demonstrates the custom strings in cells and saves the workbook.
// Keywords: Aspose.Cells | Custom GlobalizationSettings | C# | .NET | override GetBooleanValueString | override GetErrorValueString | localize boolean values | custom Excel error messages | set workbook globalization before worksheet access | Excel localization Aspose
// Common Searches: how to use custom GlobalizationSettings with Aspose.Cells | set workbook globalization settings C# Aspose | override boolean display text Aspose.Cells | custom error message localization Aspose.Cells .NET | apply GlobalizationSettings before loading worksheets
// Developer Intent: Integrate a developer‑defined GlobalizationSettings class into an Aspose.Cells workbook so that boolean and error values are rendered with custom text from the moment the workbook is opened.
// Use Cases: Generate Excel reports with localized YES/NO strings for different languages. | Replace default Excel error codes (#DIV/0!, #N/A) with user‑friendly messages in a specific market. | Guarantee consistent localization across all worksheets by setting GlobalizationSettings prior to any cell operations.
// AI Prompts: Write C# code that defines a CustomGlobalizationSettings class overriding GetBooleanValueString and GetErrorValueString, then assigns it to Workbook.Settings.GlobalizationSettings before any worksheet is accessed. | Explain step‑by‑step how to verify that custom boolean and error strings are applied in an Aspose.Cells workbook after setting GlobalizationSettings. | Suggest additional error codes to handle in CustomGlobalizationSettings while preserving the base implementation for unknown errors.

using System;
using System.IO;
using Aspose.Cells;

// Custom globalization settings – override methods as needed
// This example shows how to create a subclass of Aspose.Cells.GlobalizationSettings, override GetBooleanValueString and GetErrorValueString, and attach the custom instance to Workbook.Settings.GlobalizationSettings before any worksheet is read or written. The code demonstrates the custom strings in cells and saves the workbook.
class CustomGlobalizationSettings : GlobalizationSettings
{
    // Example: custom boolean strings
    public override string GetBooleanValueString(bool value)
    {
        return value ? "YES_CUSTOM" : "NO_CUSTOM";
    }

    // Example: custom error strings
    public override string GetErrorValueString(string err)
    {
        // Translate a few common errors, otherwise fallback to base implementation
        return err switch
        {
            "#DIV/0!" => "#DIV/0_CUSTOM",
            "#N/A"    => "#N/A_CUSTOM",
            _         => base.GetErrorValueString(err)
        };
    }
}

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputFile.xlsx";
            const string outputPath = "OutputFile.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Assign the custom globalization settings before accessing any worksheet data
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Demonstrate that the settings are applied (e.g., boolean and error cells)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(true);               // Will display "YES_CUSTOM"
            sheet.Cells["A2"].PutValue(false);              // Will display "NO_CUSTOM"
            sheet.Cells["A3"].PutValue("#DIV/0!");          // Will display "#DIV/0_CUSTOM"

            // Save the workbook with the applied settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
