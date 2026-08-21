// Title: Localize Boolean Text in Aspose.Cells by Overriding GetBooleanValueString
// Description: Demonstrates how to create a CustomBooleanGlobalizationSettings class that inherits from GlobalizationSettings, overrides GetBooleanValueString to return locale‑specific strings (e.g., French "Vrai"/"Faux"), applies the settings to a Workbook, formats cells with a BOOLEAN style, and saves the file.
// Keywords: Aspose.Cells | C# | GlobalizationSettings | GetBooleanValueString | boolean localization | French boolean strings | custom globalization | locale specific true false | Excel boolean formatting | Aspose.Cells example
// Common Searches: Aspose.Cells override GetBooleanValueString | localize true false in Excel using Aspose | custom GlobalizationSettings C# | display French boolean values in workbook | how to change boolean text in Aspose.Cells
// Developer Intent: Implement a custom GlobalizationSettings class that provides language‑specific true/false strings and attach it to a workbook so cells show localized boolean text.
// Use Cases: Show French "Vrai"/"Faux" in worksheet cells by assigning CustomBooleanGlobalizationSettings to the workbook. | Extend the overridden method to support additional languages such as German, Spanish, or Japanese. | Combine the custom settings with a BOOLEAN custom style to force text display instead of default boolean formatting.
// AI Prompts: Create a CustomBooleanGlobalizationSettings class that returns English, French, and German boolean strings and demonstrate its use with Aspose.Cells. | Explain how to apply custom globalization settings to an existing workbook and ensure cells render the localized boolean text using a BOOLEAN style. | Write unit tests for GetBooleanValueString covering at least three locales and verify the correct strings are returned.

using System;
using Aspose.Cells;

// Custom globalization settings that return locale‑specific boolean strings
// Demonstrates how to create a CustomBooleanGlobalizationSettings class that inherits from GlobalizationSettings, overrides GetBooleanValueString to return locale‑specific strings (e.g., French "Vrai"/"Faux"), applies the settings to a Workbook, formats cells with a BOOLEAN style, and saves the file.
public class CustomBooleanGlobalizationSettings : GlobalizationSettings
{
    private readonly string _locale;

    public CustomBooleanGlobalizationSettings(string locale)
    {
        _locale = locale;
    }

    // Override to provide localized true/false representations
    public override string GetBooleanValueString(bool bv)
    {
        // Example: English (default) and French
        if (_locale.Equals("fr", StringComparison.OrdinalIgnoreCase))
        {
            return bv ? "Vrai" : "Faux";
        }

        // Add more locales as needed
        return bv ? "True" : "False";
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Select the desired locale (e.g., French)
        string locale = "fr";

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = new CustomBooleanGlobalizationSettings(locale);

        Worksheet sheet = workbook.Worksheets[0];

        // Insert boolean values into cells
        sheet.Cells["A1"].PutValue(true);
        sheet.Cells["A2"].PutValue(false);

        // Set cell style to display boolean values as text
        Style boolStyle = workbook.CreateStyle();
        boolStyle.Custom = "BOOLEAN";
        sheet.Cells["A1"].SetStyle(boolStyle);
        sheet.Cells["A2"].SetStyle(boolStyle);

        // Demonstrate the overridden GetBooleanValueString method
        var gs = workbook.Settings.GlobalizationSettings;
        Console.WriteLine(gs.GetBooleanValueString(true));   // Outputs "Vrai"
        Console.WriteLine(gs.GetBooleanValueString(false));  // Outputs "Faux"

        // Save the workbook
        workbook.Save("LocalizedBooleanDemo.xlsx");
    }
}
