// Title: Aspose.Cells .NET – Generate a Workbook Localization Report with Custom Globalization Settings and Captured Error Tokens
// Description: C# sample that batch‑processes Excel files using Aspose.Cells, applies a custom ReportingGlobalizationSettings class to record every unique error token (e.g., #DIV/0!), saves a processed copy of each workbook, and creates a plain‑text report showing the full file path, the applied locale description, and any localization errors detected during cell evaluation.
// Keywords: Aspose.Cells | C# | .NET | globalization settings | custom globalization | error token capture | localization report | batch workbook processing | Excel error strings | i18n | internationalization | locale audit
// Common Searches: Aspose.Cells capture Excel error values | how to generate localization report for multiple workbooks .NET | custom globalization settings Aspose.Cells example | record #DIV/0! tokens in Excel using Aspose | batch process Excel files with locale reporting
// Developer Intent: Produce a concise report that lists each processed workbook, the custom locale applied, and any unique error strings encountered while evaluating cell values.
// Use Cases: Audit a collection of spreadsheets to identify all distinct error tokens before distribution. | Verify that a specific globalization configuration is consistently applied across a batch of Excel files. | Create processed copies of workbooks while documenting localization issues for quality‑assurance teams.
// AI Prompts: Generate a CSV version of the localization summary that includes workbook name, culture code, and error tokens. | Extend the sample to write the actual CultureInfo name (e.g., en‑US, fr‑FR) instead of a static description in the report. | Add cell address tracking so the report lists each error token together with its location (e.g., A1 – #DIV/0!).

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Custom globalization settings that records any error strings encountered
// C# sample that batch‑processes Excel files using Aspose.Cells, applies a custom ReportingGlobalizationSettings class to record every unique error token (e.g., #DIV/0!), saves a processed copy of each workbook, and creates a plain‑text report showing the full file path, the applied locale description, and any localization errors detected during cell evaluation.
public class ReportingGlobalizationSettings : GlobalizationSettings
{
    // List to store unique error strings that were requested
    public List<string> EncounteredErrors { get; } = new List<string>();

    // Override to capture error values
    public override string GetErrorValueString(string err)
    {
        // Record the original error token (e.g., "#DIV/0!")
        if (!EncounteredErrors.Contains(err))
        {
            EncounteredErrors.Add(err);
        }
        // Return the default representation (could be localized later)
        return base.GetErrorValueString(err);
    }

    // Optionally override boolean representation if needed
    public override string GetBooleanValueString(bool bv)
    {
        // Use default behavior
        return base.GetBooleanValueString(bv);
    }
}

public class WorkbookLocalizationReport
{
    // Holds report entries for each processed workbook
    private class ReportEntry
    {
        public string WorkbookPath { get; set; }
        public string LocaleDescription { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }

    // Main method to generate the report
    public static void GenerateReport(string[] workbookPaths, string reportOutputPath)
    {
        var reportEntries = new List<ReportEntry>();

        foreach (var path in workbookPaths)
        {
            // Verify the workbook file exists
            if (!File.Exists(path))
            {
                Console.WriteLine($"Warning: Workbook not found – skipping: {path}");
                continue;
            }

            try
            {
                // Load the workbook
                var workbook = new Workbook(path);

                // Create and assign custom globalization settings
                var customSettings = new ReportingGlobalizationSettings();
                workbook.Settings.GlobalizationSettings = customSettings;

                // Force evaluation of all cells to trigger error string conversion
                var cells = workbook.Worksheets[0].Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Access StringValue which uses the globalization settings
                        var _ = cells[row, col].StringValue;
                    }
                }

                // Prepare report entry
                var entry = new ReportEntry
                {
                    WorkbookPath = Path.GetFullPath(path),
                    LocaleDescription = "CustomReportingGlobalizationSettings",
                    Errors = new List<string>(customSettings.EncounteredErrors)
                };
                reportEntries.Add(entry);

                // Save a processed copy (suffix added)
                string processedPath = Path.Combine(
                    Path.GetDirectoryName(path) ?? string.Empty,
                    Path.GetFileNameWithoutExtension(path) + "_processed" + Path.GetExtension(path));

                // Ensure the target directory exists
                string processedDir = Path.GetDirectoryName(processedPath);
                if (!string.IsNullOrEmpty(processedDir) && !Directory.Exists(processedDir))
                {
                    Directory.CreateDirectory(processedDir);
                }

                workbook.Save(processedPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook '{path}': {ex.Message}");
            }
        }

        try
        {
            // Ensure the report directory exists
            string reportDir = Path.GetDirectoryName(reportOutputPath);
            if (!string.IsNullOrEmpty(reportDir) && !Directory.Exists(reportDir))
            {
                Directory.CreateDirectory(reportDir);
            }

            // Build the textual report
            using (var writer = new StreamWriter(reportOutputPath))
            {
                writer.WriteLine("Workbook Localization Report");
                writer.WriteLine("============================");
                writer.WriteLine();

                foreach (var entry in reportEntries)
                {
                    writer.WriteLine($"Workbook: {entry.WorkbookPath}");
                    writer.WriteLine($"Applied Locale: {entry.LocaleDescription}");

                    if (entry.Errors.Count == 0)
                    {
                        writer.WriteLine("Localization Errors: None");
                    }
                    else
                    {
                        writer.WriteLine("Localization Errors Encountered:");
                        foreach (var err in entry.Errors)
                        {
                            writer.WriteLine($"  - {err}");
                        }
                    }

                    writer.WriteLine(); // blank line between entries
                }
            }

            Console.WriteLine($"Report generated at: {Path.GetFullPath(reportOutputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write report: {ex.Message}");
        }
    }

    // Example usage
    public static void Main()
    {
        // Define the workbooks to process
        string[] workbooks = new string[]
        {
            "Sample1.xlsx",
            "Sample2.xlsx"
            // Add more workbook file paths as needed
        };

        // Define where the report will be saved
        string reportPath = "LocalizationReport.txt";

        // Generate the report
        GenerateReport(workbooks, reportPath);
    }
}
