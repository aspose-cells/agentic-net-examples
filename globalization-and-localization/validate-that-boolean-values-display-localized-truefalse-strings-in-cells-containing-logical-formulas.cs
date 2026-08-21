// Title: Aspose.Cells .NET – Validate Custom Boolean Localization for Formula Results
// Description: Shows how to use SettableGlobalizationSettings to replace the default TRUE/FALSE with custom strings (e.g., YES/NO), apply the settings to a workbook, evaluate a logical formula, and confirm that the cell's displayed StringValue matches the localized representation before saving the file.
// Keywords: Aspose.Cells | C# boolean localization | SettableGlobalizationSettings | custom TRUE FALSE strings | Excel boolean display | localized YES NO | formula result localization | globalization settings .NET | validate boolean string | Excel workbook localization
// Common Searches: Aspose.Cells custom boolean display | SettableGlobalizationSettings example C# | show YES instead of TRUE in Excel using Aspose | validate localized boolean string in C# | Excel boolean localization .NET
// Developer Intent: Ensure that a cell containing a logical formula shows the custom localized YES/NO strings rather than the default TRUE/FALSE.
// Use Cases: Generate region‑specific Excel reports where boolean values appear in the local language. | Automated unit test that verifies the StringValue of a formula cell equals the expected localized text. | Create documentation or dashboards with language‑appropriate boolean labels for end‑users.
// AI Prompts: Write C# code that sets custom boolean strings in Aspose.Cells, evaluates a formula, and checks that the cell's displayed string matches the localization. | Explain how SettableGlobalizationSettings affects BoolValue and StringValue for cells with logical formulas in Aspose.Cells. | Show how to retrieve and compare localized boolean strings for both true and false outcomes after workbook calculation.

using System;
using Aspose.Cells;

namespace AsposeCellsBooleanLocalizationDemo
{
    // Shows how to use SettableGlobalizationSettings to replace the default TRUE/FALSE with custom strings (e.g., YES/NO), apply the settings to a workbook, evaluate a logical formula, and confirm that the cell's displayed StringValue matches the localized representation before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create custom globalization settings and define localized boolean strings
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
            globalization.SetBooleanValueString(true, "YES");   // localized representation for TRUE
            globalization.SetBooleanValueString(false, "NO");   // localized representation for FALSE

            // Apply the custom settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // Insert a logical formula that evaluates to TRUE
            Cell boolCell = worksheet.Cells["A1"];
            boolCell.Formula = "=2>1";

            // Calculate formulas so that the result is stored in the cell
            workbook.CalculateFormula();

            // Retrieve the boolean result directly
            bool result = boolCell.BoolValue; // should be true

            // Get the localized display string via the globalization settings
            string localizedString = globalization.GetBooleanValueString(result);

            // Also retrieve the string value as shown in the cell (should reflect localization)
            string cellStringValue = boolCell.StringValue;

            // Output verification results
            Console.WriteLine($"Cell A1 formula result (BoolValue): {result}");
            Console.WriteLine($"Localized string from GetBooleanValueString: {localizedString}");
            Console.WriteLine($"Cell A1 displayed string value: {cellStringValue}");

            // Validate that the cell's displayed string matches the localized string
            if (cellStringValue == localizedString)
            {
                Console.WriteLine("Validation succeeded: Boolean value is displayed with the localized string.");
            }
            else
            {
                Console.WriteLine("Validation failed: Displayed string does not match the localized representation.");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("BooleanLocalizationDemo.xlsx");
        }
    }
}
