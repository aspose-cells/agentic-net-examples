// Title: Export each worksheet of an Excel workbook to separate HTML files with distinct TableCssId using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export individual worksheets to separate HTML files | How to set a unique TableCssId for each worksheet when saving as HTML with Aspose.Cells | Saving Excel sheets as separate HTML pages with custom table IDs in .NET | Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly per sheet example | Sanitize worksheet names for HTML file output using Aspose.Cells
// Tags: Aspose.Cells HTML export per worksheet | custom TableCssId for Aspose.Cells HTML tables | using HtmlSaveOptions to export single worksheet | worksheet name sanitization for file paths | C# generate multiple HTML files from Excel workbook

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample loads an Excel workbook, creates an output folder, and iterates through all worksheets. For each sheet it configures HtmlSaveOptions with ExportActiveWorksheetOnly and a unique TableCssId derived from the sheet index and a sanitized name, then saves the sheet as an individual HTML file. Helper methods ensure valid file names and HTML IDs.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Input Excel file – change the path as needed.
            string inputFile = "input.xlsx";

            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' does not exist.");
                return;
            }

            // Load the workbook.
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Directory where HTML files will be saved.
            string outputDir = Path.Combine(Path.GetDirectoryName(inputFile) ?? "", "HtmlSheets");
            try
            {
                Directory.CreateDirectory(outputDir);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create output directory: {ex.Message}");
                return;
            }

            // Iterate through each worksheet and export it to a separate HTML file.
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                string sheetName = sheet.Name;

                // Create HTML save options.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export only the current worksheet.
                    ExportActiveWorksheetOnly = true,

                    // Assign a unique CSS id to the generated HTML table.
                    TableCssId = $"tbl_{i}_{SanitizeId(sheetName)}"
                };

                // Build the output file name.
                string safeSheetName = SanitizeFileName(sheetName);
                string outputFile = Path.Combine(outputDir,
                    $"{Path.GetFileNameWithoutExtension(inputFile)}_{safeSheetName}.html");

                try
                {
                    // Save the current worksheet as HTML.
                    workbook.Save(outputFile, htmlOptions);
                    Console.WriteLine($"Worksheet '{sheetName}' exported to '{outputFile}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export worksheet '{sheetName}': {ex.Message}");
                }
            }

            Console.WriteLine("Export completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Replace invalid filename characters with underscores.
    private static string SanitizeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }

    // Ensure HTML id is valid: start with a letter, contain only allowed characters.
    private static string SanitizeId(string id)
    {
        // Allowed characters: letters, digits, hyphen, underscore, colon, period.
        // Replace any other character with an underscore.
        string pattern = @"[^A-Za-z0-9\-\_:\.]";
        string result = Regex.Replace(id, pattern, "_");

        // Ensure it does not start with a digit.
        if (result.Length > 0 && char.IsDigit(result, 0))
        {
            result = "_" + result;
        }
        return result;
    }
}
