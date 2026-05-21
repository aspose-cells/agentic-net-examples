using System;
using Aspose.Cells;

namespace AsposeCellsFallbackDemo
{
    // Custom globalization settings that intentionally throws an exception
    // for demonstration purposes.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetAllName()
        {
            // Simulate a failure in custom label method.
            throw new InvalidOperationException("Custom GetAllName failed.");
        }

        // You can override other label methods similarly if needed.
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Assign the custom globalization settings.
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Attempt to retrieve a custom label.
            string allLabel;
            try
            {
                // This call will invoke the overridden method which throws.
                allLabel = workbook.Settings.GlobalizationSettings.GetAllName();
                Console.WriteLine($"Custom '(All)' label: {allLabel}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Custom label method threw an exception: {ex.Message}");
                Console.WriteLine("Reverting to default globalization settings.");

                // Revert to the default globalization settings.
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings();

                // Retry using the default implementation.
                allLabel = workbook.Settings.GlobalizationSettings.GetAllName();
                Console.WriteLine($"Default '(All)' label after fallback: {allLabel}");
            }

            // Demonstrate that the workbook can still be used normally.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B1"].Formula = $"={allLabel}(A2)"; // Example usage; may not be a valid formula.

            // Save the workbook.
            workbook.Save("FallbackDemo.xlsx");
        }
    }
}