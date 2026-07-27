using System;
using Aspose.Cells;

namespace AsposeCellsCustomGlobalization
{
    // Custom globalization settings that maps standard function names to localized names and vice‑versa.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Map standard English function names to custom localized equivalents.
        public override string GetLocalFunctionName(string standardName)
        {
            return standardName switch
            {
                "SUM" => "LOCALSUM",
                "AVERAGE" => "LOCALAVG",
                _ => base.GetLocalFunctionName(standardName)
            };
        }

        // Map localized function names back to the standard English names.
        public override string GetStandardFunctionName(string localName)
        {
            return localName switch
            {
                "LOCALSUM" => "SUM",
                "LOCALAVG" => "AVERAGE",
                _ => base.GetStandardFunctionName(localName)
            };
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Assign the custom globalization settings to the workbook.
                workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

                // Populate sample data.
                worksheet.Cells["B1"].PutValue(10);
                worksheet.Cells["B2"].PutValue(20);
                worksheet.Cells["B3"].PutValue(30);

                // Use the localized function name in a formula.
                Cell formulaCell = worksheet.Cells["B4"];
                formulaCell.Formula = "=LOCALSUM(B1:B3)";

                // Calculate all formulas in the workbook.
                workbook.CalculateFormula();

                // Retrieve the result safely.
                double result = 0;
                if (formulaCell.Value is double d)
                {
                    result = d;
                }
                else
                {
                    // Fallback: try to convert the string result.
                    double.TryParse(formulaCell.StringValue, out result);
                }

                Console.WriteLine($"Result of LOCALSUM(B1:B3): {result}");

                // Save the workbook.
                string outputPath = "CustomGlobalizationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}