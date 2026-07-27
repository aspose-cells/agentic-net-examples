using System;
using Aspose.Cells;

namespace AsposeCellsFallbackDemo
{
    // Custom globalization settings that deliberately throws an exception for demonstration.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetAllName()
        {
            // Simulate a failure in custom label retrieval.
            throw new InvalidOperationException("Custom label retrieval failed.");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided creation rule).
            Workbook workbook = new Workbook();

            // Assign custom globalization settings that may throw.
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Attempt to use a custom label method.
            try
            {
                // This call is expected to throw.
                string allLabel = ((CustomGlobalizationSettings)workbook.Settings.GlobalizationSettings).GetAllName();
                Console.WriteLine($"Custom '(All)' label: {allLabel}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Custom label method failed: {ex.Message}");
                Console.WriteLine("Reverting to default globalization settings.");

                // Revert to the default globalization settings (fallback mechanism).
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings();
            }

            // Continue working with the workbook using the (now) default settings.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].Formula = "=SUM(1,2,3)";

            // Calculate formulas to ensure everything works with default settings.
            workbook.CalculateFormula();

            // Save the workbook (uses the provided save rule).
            workbook.Save("FallbackDemoOutput.xlsx");
        }
    }
}