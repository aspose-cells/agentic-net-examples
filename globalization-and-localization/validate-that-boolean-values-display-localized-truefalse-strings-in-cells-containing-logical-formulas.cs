using System;
using Aspose.Cells;

namespace AsposeCellsBooleanLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create custom globalization settings with localized boolean strings
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
            globalization.SetBooleanValueString(true, "VRAI");   // French for TRUE
            globalization.SetBooleanValueString(false, "FAUX"); // French for FALSE

            // Apply the custom settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // Insert logical formulas that evaluate to boolean values
            worksheet.Cells["A1"].Formula = "=1<2"; // evaluates to TRUE
            worksheet.Cells["A2"].Formula = "=1>2"; // evaluates to FALSE

            // Calculate formulas so that BoolValue and StringValue are populated
            workbook.CalculateFormula();

            // Retrieve the boolean values from the cells
            bool boolA1 = worksheet.Cells["A1"].BoolValue;
            bool boolA2 = worksheet.Cells["A2"].BoolValue;

            // Get the localized display strings using the globalization settings
            string localizedA1 = globalization.GetBooleanValueString(boolA1);
            string localizedA2 = globalization.GetBooleanValueString(boolA2);

            // Also retrieve the string representation directly from the cells
            string cellStringA1 = worksheet.Cells["A1"].StringValue;
            string cellStringA2 = worksheet.Cells["A2"].StringValue;

            // Output the results for verification
            Console.WriteLine($"Cell A1 BoolValue: {boolA1}, Localized via settings: {localizedA1}, Cell StringValue: {cellStringA1}");
            Console.WriteLine($"Cell A2 BoolValue: {boolA2}, Localized via settings: {localizedA2}, Cell StringValue: {cellStringA2}");

            // Save the workbook (demonstrates that the localized strings are persisted)
            workbook.Save("BooleanLocalizationDemo.xlsx");
        }
    }
}