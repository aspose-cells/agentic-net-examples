// Title: Batch Apply Custom GlobalizationSettings to Excel Workbooks with Aspose.Cells (.NET)
// Description: C# sample that defines a CustomGlobalizationSettings class (e.g., Russian boolean and error strings) and a WorkbookBatchProcessor which scans a source folder, loads each .xls/.xlsx/.xlsb/.xlsm file, assigns the custom globalization settings, and saves the localized copy to a destination folder while preserving file names.
// Keywords: Aspose.Cells | C# | GlobalizationSettings | custom localization | batch processing | Excel folder automation | Russian error messages | boolean translation | SettableGlobalizationSettings | list separator
// Common Searches: apply custom GlobalizationSettings to multiple Excel files Aspose.Cells | batch localize Excel workbooks .NET | Aspose.Cells change error text for a folder of workbooks | set list separator for all workbooks using Aspose.Cells | C# code to process Excel files with custom globalization
// Developer Intent: Automatically assign a user‑defined GlobalizationSettings implementation to every workbook in a directory and save the localized versions.
// Use Cases: Convert a collection of financial reports to display Russian TRUE/FALSE and error texts before distribution. | Prepare Excel templates for European markets by switching the list separator to a semicolon in bulk. | Standardize error messages across legacy workbooks to meet regulatory language requirements.
// AI Prompts: Write C# code that iterates over a folder of .xlsx files, loads each workbook with Aspose.Cells, applies a GlobalizationSettings subclass returning French translations, and saves the results to an output directory. | Explain how to extend GlobalizationSettings in Aspose.Cells to provide custom boolean and error strings, then use it in a batch process for multiple workbooks. | Provide step‑by‑step instructions for changing the list separator via SettableGlobalizationSettings while processing a batch of Excel files with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchGlobalization
{
    // Custom globalization settings derived from GlobalizationSettings.
    // Override methods to provide localized strings for booleans and errors.
    // C# sample that defines a CustomGlobalizationSettings class (e.g., Russian boolean and error strings) and a WorkbookBatchProcessor which scans a source folder, loads each .xls/.xlsx/.xlsb/.xlsm file, assigns the custom globalization settings, and saves the localized copy to a destination folder while preserving file names.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool value)
        {
            // Example: Russian boolean strings.
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string err)
        {
            // Map standard error texts to localized versions.
            return err switch
            {
                "#NAME?" => "#ИМЯ?",
                "#DIV/0!" => "#ДЕЛ/0!",
                "#REF!" => "#ССЫЛКА!",
                "#VALUE!" => "#ЗНАЧ!",
                "#N/A" => "#Н/Д",
                "#NUM!" => "#ЧИСЛО!",
                "#NULL!" => "#ПУСТО!",
                _ => base.GetErrorValueString(err)
            };
        }
    }

    public static class WorkbookBatchProcessor
    {
        // Applies the custom globalization settings to every workbook in the source folder.
        // Processed files are saved to the destination folder preserving original file names.
        public static void ApplyGlobalizationToFolder(string sourceFolder, string destinationFolder)
        {
            // Ensure destination folder exists.
            Directory.CreateDirectory(destinationFolder);

            // Get all supported Excel files in the source folder.
            string[] files = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in files)
            {
                // Filter by known Excel extensions.
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsb" && ext != ".xlsm")
                    continue;

                // Load the workbook (create rule: Workbook(string) constructor).
                Workbook wb = new Workbook(filePath);

                // Apply the custom globalization settings.
                wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

                // Optionally, you can also use SettableGlobalizationSettings for additional tweaks.
                // Example: change list separator to semicolon.
                // var settable = new SettableGlobalizationSettings();
                // settable.SetListSeparator(';');
                // wb.Settings.GlobalizationSettings = settable;

                // Save the modified workbook to the destination folder (save rule: Workbook.Save(string)).
                string destPath = Path.Combine(destinationFolder, Path.GetFileName(filePath));
                wb.Save(destPath);
            }
        }
    }

    // Example usage.
    class Program
    {
        static void Main()
        {
            string inputFolder = @"C:\InputWorkbooks";
            string outputFolder = @"C:\OutputWorkbooks";

            WorkbookBatchProcessor.ApplyGlobalizationToFolder(inputFolder, outputFolder);

            Console.WriteLine("Batch processing completed.");
        }
    }
}
