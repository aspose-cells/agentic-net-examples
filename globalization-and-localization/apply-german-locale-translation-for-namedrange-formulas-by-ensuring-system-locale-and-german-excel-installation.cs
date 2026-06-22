using System;
using System.IO;
using Aspose.Cells;

class GermanLocaleNamedRangeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the workbook region to Germany to enable German locale
            workbook.Settings.Region = CountryCode.Germany;

            // Create custom globalization settings
            SettableGlobalizationSettings settings = new SettableGlobalizationSettings();

            // Map standard English function names to German equivalents (bidirectional)
            settings.SetLocalFunctionName("SUM", "SUMME", true);
            settings.SetLocalFunctionName("AVERAGE", "MITTELWERT", true);

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = settings;

            // Populate sample data in cells A1:A5
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["A3"].PutValue(30);
            worksheet.Cells["A4"].PutValue(40);
            worksheet.Cells["A5"].PutValue(50);

            // Define a named range that uses a German localized formula
            int nameIndex = workbook.Worksheets.Names.Add("TotalGerman"); // returns index
            Name totalGerman = workbook.Worksheets.Names[nameIndex];
            totalGerman.RefersTo = "=SUMME(A1:A5)";

            // Use the named range in a cell formula
            worksheet.Cells["B1"].Formula = "=TotalGerman";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the calculated result
            Console.WriteLine("Result of German localized named‑range formula: " + worksheet.Cells["B1"].Value);

            // Save the workbook
            string outputPath = "GermanLocaleNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}