// Title: Customize Excel Error Messages in Aspose.Cells .NET by Overriding GetErrorValueString
// Description: Shows how to subclass GlobalizationSettings, override GetErrorValueString to replace default Excel error codes with localized text, apply the custom settings to a Workbook, generate a #DIV/0! error, and read the tailored display string.
// Keywords: Aspose.Cells | GlobalizationSettings | GetErrorValueString | custom error messages | Excel error localization | C# | override error strings | internationalization | localized #DIV/0! | custom GlobalizationSettings class
// Common Searches: Aspose.Cells override GetErrorValueString example | how to localize Excel error messages with Aspose.Cells | custom GlobalizationSettings for error strings .NET | replace #DIV/0! with custom text in Aspose.Cells | Excel error code localization C# Aspose
// Developer Intent: Create a GlobalizationSettings subclass that returns language‑specific strings for Excel error codes.
// Use Cases: Generate workbooks that display user‑friendly error text for non‑technical audiences. | Provide localized error descriptions in multinational reporting solutions. | Maintain consistent error wording across a suite of automated Excel exports.
// AI Prompts: Write a CustomErrorGlobalizationSettings class that maps #REF! and #NUM! to German translations. | Demonstrate applying the custom globalization settings to multiple workbooks in a batch loop. | Create unit tests that verify each overridden error code returns the expected localized string.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom globalization settings that provide localized error messages
    // Shows how to subclass GlobalizationSettings, override GetErrorValueString to replace default Excel error codes with localized text, apply the custom settings to a Workbook, generate a #DIV/0! error, and read the tailored display string.
    public class CustomErrorGlobalizationSettings : GlobalizationSettings
    {
        // Override GetErrorValueString to map default error strings to custom ones
        public override string GetErrorValueString(string err)
        {
            // Map specific Excel error codes to localized messages
            return err switch
            {
                "#DIV/0!" => "Custom Division Error",
                "#VALUE!" => "Custom Type Mismatch",
                "#NAME?" => "Custom Identifier Error",
                "#N/A"    => "Custom Not Available",
                _ => base.GetErrorValueString(err) // Fallback to default behavior
            };
        }
    }

    public class GlobalizationSettingsMethodGetErrorValueStringDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomErrorGlobalizationSettings();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a cell that will generate a division by zero error
            Cell errorCell = worksheet.Cells["A1"];
            errorCell.Formula = "=1/0"; // This will produce #DIV/0! error

            // Calculate formulas to evaluate the error
            workbook.CalculateFormula();

            // Retrieve the display string for the error cell (uses our custom mapping)
            string errorDisplay = errorCell.DisplayStringValue;

            // Output the custom error message to the console
            Console.WriteLine($"Error display value: {errorDisplay}");

            // Save the workbook to verify the result
            workbook.Save("MethodGetErrorValueStringDemo.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            GlobalizationSettingsMethodGetErrorValueStringDemo.Run();
        }
    }
}
