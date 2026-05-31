using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
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
            // bidirectional = true enables automatic reverse mapping (local -> standard)
            settings.SetLocalFunctionName("SUM", "SOMME", true);

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = settings;

            // Verify the mapping by retrieving the localized name for "SUM"
            string localizedName = settings.GetLocalFunctionName("SUM");
            Console.WriteLine($"Localized name for 'SUM' is: {localizedName}");

            // Populate sample data in column B (B1:B5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[$"B{i + 1}"].PutValue(i + 1); // Values 1,2,3,4,5
            }

            // Use the localized function name in a formula
            sheet.Cells["A1"].Formula = $"={localizedName}(B1:B5)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the calculation result
            Console.WriteLine($"Result of formula in A1 ({sheet.Cells["A1"].Formula}): {sheet.Cells["A1"].Value}");

            // Save the workbook – when opened in Excel (with French locale) the formula will appear as =SOMME(...)
            workbook.Save("LocalizedFunctionDemo.xlsx");
        }
    }
}