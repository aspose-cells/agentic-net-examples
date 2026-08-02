using System;
using Aspose.Cells;

class VerifyLocalizedFunction
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create globalization settings and map the standard "SUM" function to the French localized name "SOMME"
        SettableGlobalizationSettings settings = new SettableGlobalizationSettings();
        settings.SetLocalFunctionName("SUM", "SOMME", true);

        // Apply the settings to the workbook
        workbook.Settings.GlobalizationSettings = settings;

        // Verify the mapping using GetLocalFunctionName
        string localizedName = settings.GetLocalFunctionName("SUM");
        Console.WriteLine($"Localized name for 'SUM' is: {localizedName}");

        // Populate some sample data
        worksheet.Cells["B1"].PutValue(10);
        worksheet.Cells["B2"].PutValue(20);
        worksheet.Cells["B3"].PutValue(30);

        // Use the localized function name in a formula
        worksheet.Cells["A1"].Formula = "=SOMME(B1:B3)";

        // Calculate formulas
        workbook.CalculateFormula();

        // Output the calculation result
        Console.WriteLine($"Result of localized formula in A1: {worksheet.Cells["A1"].Value}");

        // Save the workbook – when opened in Excel with a matching locale, the formula will be recognized
        workbook.Save("LocalizedFunctionDemo.xlsx");
    }
}