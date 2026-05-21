using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchGlobalization
{
    // Custom batch processor that applies globalization settings to every workbook in a folder
    public static class GlobalizationBatchProcessor
    {
        // Entry point
        public static void ProcessFolder(string inputFolder, string outputFolder)
        {
            // Ensure output folder exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all Excel files (you can adjust the pattern as needed)
            string[] files = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                // Load the workbook (uses the provided load rule)
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Create and configure custom globalization settings (uses the provided create rule)
                    SettableGlobalizationSettings gSettings = new SettableGlobalizationSettings();

                    // Example customizations
                    gSettings.SetListSeparator(';');                     // Use semicolon as list separator
                    gSettings.SetBooleanValueString(true, "TRUE_CUSTOM");   // Custom true string
                    gSettings.SetBooleanValueString(false, "FALSE_CUSTOM"); // Custom false string
                    gSettings.SetLocalFunctionName("SUM", "SUMME", true);   // Localized SUM function (German)
                    gSettings.SetLocalFunctionName("AVERAGE", "MITTELWERT", true); // Localized AVERAGE

                    // Apply the settings to the workbook
                    workbook.Settings.GlobalizationSettings = gSettings;

                    // Determine output path (same name, different folder)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                    // Save the modified workbook (uses the provided save rule)
                    workbook.Save(outputPath);
                }
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string sourceFolder = @"C:\InputWorkbooks";
            string destinationFolder = @"C:\OutputWorkbooks";

            GlobalizationBatchProcessor.ProcessFolder(sourceFolder, destinationFolder);

            Console.WriteLine("Batch processing completed.");
        }
    }
}