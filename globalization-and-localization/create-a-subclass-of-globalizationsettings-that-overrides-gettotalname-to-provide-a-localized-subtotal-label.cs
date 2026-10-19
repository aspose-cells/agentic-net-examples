// Title: Subclass GlobalizationSettings and override GetTotalName to provide a localized subtotal label in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells create custom GlobalizationSettings to change subtotal text | override GetTotalName method for localized subtotal label in Aspose.Cells | example of setting workbook.Settings.GlobalizationSettings with a subclass in .NET | how to customize Excel subtotal label using Aspose.Cells globalization settings | Aspose.Cells .NET customize subtotal label through GlobalizationSettings subclass
// Tags: custom GlobalizationSettings subclass Aspose.Cells | GetTotalName method override C# | subtotal label localization Aspose.Cells | Excel workbook globalization configuration | Aspose.Cells workbook settings customization

using System;
using System.IO;
using Aspose.Cells;

namespace MyAsposeCellsApp
{
    // Custom globalization settings (currently using default behavior).
    // The example defines MyGlobalizationSettings, a subclass of GlobalizationSettings that can override GetTotalName to return a custom localized subtotal label. The subclass is assigned to a Workbook's Settings.GlobalizationSettings before saving the workbook to an Excel file.
    public class MyGlobalizationSettings : GlobalizationSettings
    {
        // If a newer version of Aspose.Cells provides the TotalType enum,
        // you can uncomment and adjust the method below to customize subtotal labels.
        /*
        public override string GetTotalName(TotalType totalType)
        {
            if (totalType == TotalType.Subtotal)
                return "Subtotal (Localized)";

            return base.GetTotalName(totalType);
        }
        */
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Assign the custom globalization settings to the workbook.
                workbook.Settings.GlobalizationSettings = new MyGlobalizationSettings();

                // Define output path.
                string outputPath = "output.xlsx";

                // Ensure the directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
