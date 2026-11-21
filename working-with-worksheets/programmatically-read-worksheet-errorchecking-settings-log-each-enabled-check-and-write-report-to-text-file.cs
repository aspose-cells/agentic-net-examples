// Title: Extract Excel worksheet error‑checking options with Aspose.Cells for .NET and generate a text report
// Common Searches: Aspose.Cells C# read worksheet error checking settings and export to txt | How to get ErrorCheckingOptions from a workbook using Aspose.Cells .NET | Generate a report of Excel error‑checking flags with Aspose.Cells in C# | Retrieve and log Excel error checking configuration via Aspose.Cells reflection | Save Aspose.Cells worksheet validation settings to a text file
// Tags: Aspose.Cells error checking options extraction | C# extract Excel worksheet error checking flags | Aspose.Cells Settings.ErrorCheckingOptions reflection access | text file report of Excel error checking status | audit Excel error checking configuration with Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The sample loads an Excel workbook using Aspose.Cells for .NET, accesses the Settings.ErrorCheckingOptions object via reflection to handle version differences, enumerates each error‑checking flag (formula, number stored as text, inconsistent formula, etc.), builds a formatted text report indicating enabled or disabled status, and writes the report to a .txt file while gracefully handling unsupported library versions.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths.
            const string inputPath = "input.xlsx";
            const string outputPath = "ErrorCheckingReport.txt";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Attempt to obtain error‑checking options via reflection (API may differ between versions).
            dynamic errorOptions = GetErrorCheckingOptions(workbook);

            // Prepare a StringBuilder for the report.
            StringBuilder report = new StringBuilder();

            report.AppendLine("Worksheet Error‑Checking Settings Report");
            report.AppendLine($"Generated on: {DateTime.Now}");
            report.AppendLine();

            if (errorOptions != null)
            {
                // Log each error‑checking option if available.
                report.AppendLine($"Formula error checking: {(errorOptions.IsFormulaErrorCheckingEnabled ? "Enabled" : "Disabled")}");
                report.AppendLine($"Number stored as text error checking: {(errorOptions.IsNumberStoredAsTextErrorCheckingEnabled ? "Enabled" : "Disabled")}");
                report.AppendLine($"Inconsistent formula error checking: {(errorOptions.IsInconsistentFormulaErrorCheckingEnabled ? "Enabled" : "Disabled")}");
                report.AppendLine($"Empty cell reference error checking: {(errorOptions.IsEmptyCellReferenceErrorCheckingEnabled ? "Enabled" : "Disabled")}");
                report.AppendLine($"List data validation error checking: {(errorOptions.IsListDataValidationErrorCheckingEnabled ? "Enabled" : "Disabled")}");
                report.AppendLine($"Unlocked formula error checking: {(errorOptions.IsUnlockedFormulaErrorCheckingEnabled ? "Enabled" : "Disabled")}");
                report.AppendLine($"Smart tag error checking: {(errorOptions.IsSmartTagErrorCheckingEnabled ? "Enabled" : "Disabled")}");
                report.AppendLine($"Background error checking: {(errorOptions.IsBackgroundErrorCheckingEnabled ? "Enabled" : "Disabled")}");
            }
            else
            {
                report.AppendLine("Error checking options are not supported in this Aspose.Cells version.");
            }

            // Write the report to the output file.
            File.WriteAllText(outputPath, report.ToString());

            Console.WriteLine($"Report generated successfully at: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Retrieves the ErrorCheckingOptions object via reflection to handle version differences.
    private static dynamic GetErrorCheckingOptions(Workbook workbook)
    {
        try
        {
            var settingsProp = workbook.GetType().GetProperty("Settings");
            if (settingsProp == null) return null;

            var settingsObj = settingsProp.GetValue(workbook);
            if (settingsObj == null) return null;

            var errorProp = settingsObj.GetType().GetProperty("ErrorCheckingOptions");
            if (errorProp == null) return null;

            return errorProp.GetValue(settingsObj);
        }
        catch
        {
            // If any reflection step fails, return null indicating unsupported feature.
            return null;
        }
    }
}
