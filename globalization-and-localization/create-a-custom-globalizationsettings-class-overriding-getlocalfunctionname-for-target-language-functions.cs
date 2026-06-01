using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomGlobalizationDemo
{
    // Custom globalization settings that maps standard function names to localized ones.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override to provide locale‑dependent function names.
        public override string GetLocalFunctionName(string standardName)
        {
            // Map the standard "SUM" function to a custom localized name "LOCALSUM".
            if (standardName.Equals("SUM", StringComparison.OrdinalIgnoreCase))
                return "LOCALSUM";

            // Map the standard "AVERAGE" function to a custom localized name "LOCALAVERAGE".
            if (standardName.Equals("AVERAGE", StringComparison.OrdinalIgnoreCase))
                return "LOCALAVERAGE";

            // For all other functions fall back to the base implementation.
            return base.GetLocalFunctionName(standardName);
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Assign the custom globalization settings to the workbook.
                workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

                // NOTE: In recent Aspose.Cells versions, function localization is enabled by default.
                // If using an older version where explicit enabling is required, uncomment the following line:
                // workbook.Settings.CalcEngineSettings.EnableFunctionLocalization = true;

                // Populate sample data that will be summed.
                worksheet.Cells["B1"].PutValue(5);
                worksheet.Cells["B2"].PutValue(15);
                worksheet.Cells["B3"].PutValue(25);

                // Use the localized function name in a formula.
                Cell formulaCell = worksheet.Cells["B4"];
                formulaCell.Formula = "=LOCALSUM(B1:B3)";

                // Calculate the formula.
                workbook.CalculateFormula();

                // Output the result to the console.
                Console.WriteLine($"Result of LOCALSUM(B1:B3): {formulaCell.DoubleValue}");

                // Save the workbook to a file.
                string outputPath = "CustomGlobalizationDemo.xlsx";

                // Ensure the directory exists before saving.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

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