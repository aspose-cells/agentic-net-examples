using System;
using System.IO;
using Aspose.Cells;
using System.Globalization;

namespace AsposeCellsBatchGlobalization
{
    // Custom globalization settings derived from GlobalizationSettings.
    // Override methods to provide localized strings for booleans and errors.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool value)
        {
            // Example: Russian boolean strings.
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string error)
        {
            // Map standard error texts to localized versions.
            return error switch
            {
                "#NAME?" => "#ИМЯ?",
                "#DIV/0!" => "#ДЕЛ/0!",
                "#REF!" => "#ССЫЛКА!",
                "#VALUE!" => "#ЗНАЧ!",
                "#N/A" => "#Н/Д",
                "#NUM!" => "#ЧИСЛО!",
                "#NULL!" => "#ПУСТО!",
                _ => base.GetErrorValueString(error)
            };
        }
    }

    public static class WorkbookBatchProcessor
    {
        // Processes all Excel files in the specified input folder,
        // applies the custom globalization settings, and saves them to the output folder.
        public static void ApplyGlobalizationToFolder(string inputFolder, string outputFolder)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Supported Excel extensions.
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsb", ".xlsm", ".ods" };

            foreach (string filePath in Directory.EnumerateFiles(inputFolder))
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue; // Skip non‑Excel files.

                // Load the workbook using the constructor that accepts a file path.
                Workbook workbook = new Workbook(filePath);

                // Apply the custom globalization settings.
                workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

                // Build the output file path (preserve original file name).
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the workbook, overwriting if it already exists.
                workbook.Save(outputPath);
            }
        }
    }

    // Example usage.
    class Program
    {
        static void Main()
        {
            // Folder containing source workbooks.
            string sourceFolder = @"C:\InputWorkbooks";

            // Folder where processed workbooks will be saved.
            string destinationFolder = @"C:\OutputWorkbooks";

            WorkbookBatchProcessor.ApplyGlobalizationToFolder(sourceFolder, destinationFolder);

            Console.WriteLine("Batch processing completed.");
        }
    }
}