using Aspose.Cells;
using System;

class BooleanLocalizationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Create custom globalization settings and define localized boolean strings
        SettableGlobalizationSettings gSettings = new SettableGlobalizationSettings();
        gSettings.SetBooleanValueString(true, "YES_LOCAL");
        gSettings.SetBooleanValueString(false, "NO_LOCAL");

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = gSettings;

        // Insert logical formulas that evaluate to boolean values
        sheet.Cells["A1"].Formula = "=2>1"; // evaluates to true
        sheet.Cells["A2"].Formula = "=1>2"; // evaluates to false

        // Calculate the formulas so that the cells contain the evaluated results
        workbook.CalculateFormula();

        // Retrieve the displayed string values from the cells
        string displayedA1 = sheet.Cells["A1"].StringValue; // should show "YES_LOCAL"
        string displayedA2 = sheet.Cells["A2"].StringValue; // should show "NO_LOCAL"

        // Use GetBooleanValueString to obtain the expected localized strings
        string expectedTrue = gSettings.GetBooleanValueString(true);
        string expectedFalse = gSettings.GetBooleanValueString(false);

        // Output the results and validation status
        Console.WriteLine($"Cell A1 displays: {displayedA1} (expected: {expectedTrue})");
        Console.WriteLine($"Cell A2 displays: {displayedA2} (expected: {expectedFalse})");

        bool isValid = displayedA1 == expectedTrue && displayedA2 == expectedFalse;
        Console.WriteLine("Localization validation " + (isValid ? "passed" : "failed"));

        // Save the workbook
        workbook.Save("BooleanLocalizationDemo.xlsx");
    }
}