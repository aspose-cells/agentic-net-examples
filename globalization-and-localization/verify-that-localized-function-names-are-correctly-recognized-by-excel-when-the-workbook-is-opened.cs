// Title: Verify French Excel Function Mapping (SUM → SOMME) with Aspose.Cells for .NET
// Description: Shows how to configure Aspose.Cells SettableGlobalizationSettings to map the English function SUM to the French name SOMME, insert sample data, apply the localized formula, evaluate it, and save a workbook that Excel presents with French function names.
// Keywords: Aspose.Cells | .NET | Excel localized functions | French function name mapping | SettableGlobalizationSettings | globalization settings | bidirectional function mapping | formula localization | workbook calculation
// Common Searches: map English Excel function to French using Aspose.Cells | verify localized formula names in a .NET workbook | set custom globalization settings for Excel functions | does Excel show French function names after saving with Aspose.Cells | Aspose.Cells French SUM SOMME example
// Developer Intent: Ensure that a custom mapping of the English function SUM to the French equivalent SOMME is applied, calculated correctly in code, and recognized by Excel when the file is opened.
// Use Cases: Automatically generate workbooks that display French function names for francophone users. | Create template files with locale‑specific formulas before distribution. | Programmatically test that formulas using localized names evaluate correctly prior to saving.
// AI Prompts: Provide an example that maps several English Excel functions to their German equivalents with SettableGlobalizationSettings. | Show code to read back a localized function name from a saved workbook to confirm the mapping. | Explain how to enable and use bidirectional function name mapping for custom locales in Aspose.Cells.

using System;
using Aspose.Cells;

namespace LocalizedFunctionVerification
{
    // Shows how to configure Aspose.Cells SettableGlobalizationSettings to map the English function SUM to the French name SOMME, insert sample data, apply the localized formula, evaluate it, and save a workbook that Excel presents with French function names.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Create customizable globalization settings
            SettableGlobalizationSettings settings = new SettableGlobalizationSettings();

            // Map the standard English function name "SUM" to the French localized name "SOMME"
            // Bidirectional = true enables automatic reverse mapping (local -> standard)
            settings.SetLocalFunctionName("SUM", "SOMME", true);

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = settings;

            // Verify the mapping by retrieving the localized name via GetLocalFunctionName
            string localizedName = settings.GetLocalFunctionName("SUM");
            Console.WriteLine($"Localized name for 'SUM' is: {localizedName}");

            // Populate sample data in column B (B1:B5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[$"B{i + 1}"].PutValue(i + 1); // Values 1,2,3,4,5
            }

            // Use the localized function name in a formula
            sheet.Cells["A1"].Formula = $"={localizedName}(B1:B5)";

            // Calculate the formula
            workbook.CalculateFormula();

            // Output the calculation result
            Console.WriteLine($"Result of formula using localized function: {sheet.Cells["A1"].Value}");

            // Save the workbook – Excel will display the formula using the locale of the user opening the file
            workbook.Save("LocalizedFunctionDemo.xlsx");
        }
    }
}
