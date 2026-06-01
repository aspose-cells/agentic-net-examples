using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Apply globalization settings that use standard (en‑US) function names
            SettableGlobalizationSettings gSettings = new SettableGlobalizationSettings();
            workbook.Settings.GlobalizationSettings = gSettings;

            // Optional: force workbook region to United States.
            // Uncomment the line below if the CountryCode enum in your Aspose.Cells version contains UnitedStates.
            // workbook.Settings.Region = CountryCode.UnitedStates;

            // Sample data for the formula
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(15);

            // Configure formula parsing options: treat formula as English (LocaleDependent = false)
            FormulaParseOptions options = new FormulaParseOptions
            {
                LocaleDependent = false,
                R1C1Style = false
            };

            // Set a formula using the English function name "SUM"
            sheet.Cells["A3"].SetFormula("=SUM(A1:A2)", options);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the calculated result
            Console.WriteLine("Result of SUM(A1:A2): " + sheet.Cells["A3"].Value);

            // Save the workbook
            string outputPath = "FormulaEnglishFunction.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}