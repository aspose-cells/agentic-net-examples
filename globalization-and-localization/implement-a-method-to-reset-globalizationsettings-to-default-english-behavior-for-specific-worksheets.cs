// Title: Reset Aspose.Cells workbook globalization to English (en-US) in C# for selected worksheets
// Common Searches: Aspose.Cells set workbook culture to en-US in C# | reset globalization settings for specific worksheets using Aspose.Cells .NET | verify worksheet names before applying workbook-wide locale Aspose.Cells | change Excel file locale to English with Aspose.Cells API | C# example for workbook.Settings.CultureInfo Aspose.Cells
// Tags: Aspose.Cells workbook cultureinfo en-US | C# reset workbook globalization Aspose.Cells | validate worksheet existence Aspose.Cells | Excel file locale change Aspose.Cells .NET | set workbook-wide culture Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

// The ResetGlobalizationSettings method sets Workbook.Settings.CultureInfo to "en-US" and iterates over a list of worksheet names to confirm they exist, noting that Aspose.Cells applies globalization only at the workbook level. The example loads a workbook, validates the target sheets, applies the English locale, and saves the file with proper error handling.
public static class GlobalizationHelper
{
    // Resets the globalization settings of the workbook to English (United States).
    // Aspose.Cells applies these settings at the workbook level; per‑worksheet settings are not available.
    public static void ResetGlobalizationSettings(Workbook workbook, IEnumerable<string> worksheetNames)
    {
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        // Set workbook‑wide culture to English (United States).
        workbook.Settings.CultureInfo = new CultureInfo("en-US");
        // The UseEnglishFunctionName property may not be available in all versions; culture setting is sufficient.

        // Verify that the requested worksheets exist (optional, no per‑sheet changes possible).
        foreach (var name in worksheetNames)
        {
            bool exists = workbook.Worksheets.Any(ws => ws.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (!exists)
                continue; // Skip missing sheets.
            // No additional per‑sheet globalization actions required.
        }
    }

    // Example usage demonstrating load, reset, and save with safety checks.
    public static void Example()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists before attempting to load.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the existing workbook.
            var wb = new Workbook(inputPath);

            // Reset globalization settings for specific worksheets (validation only).
            ResetGlobalizationSettings(wb, new List<string> { "Sheet1", "Data" });

            // Save the modified workbook.
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed.
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    // Entry point required for the console application.
    public static void Main(string[] args)
    {
        GlobalizationHelper.Example();
    }
}
