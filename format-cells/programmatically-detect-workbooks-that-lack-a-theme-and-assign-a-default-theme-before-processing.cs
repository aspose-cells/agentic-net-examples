// Title: Apply a Default Theme to Excel Workbooks Missing a Theme with Aspose.Cells for .NET
// Description: C# code that scans a list of Excel files, loads each workbook with Aspose.Cells, detects a missing Theme, creates a temporary workbook containing the built‑in default theme, copies that theme via Workbook.CopyTheme, and saves the updated file to an output folder.
// Keywords: Aspose.Cells default theme | C# detect missing workbook theme | CopyTheme Aspose.Cells | apply Excel theme programmatically | Workbook.Theme check .NET | set default Excel theme | Aspose.Cells theme handling
// Common Searches: how to set a default theme for Excel files using Aspose.Cells | detect workbooks without a theme in C# | copy theme from a template workbook Aspose.Cells | Workbook.CopyTheme example .NET | Aspose.Cells check if workbook theme is empty
// Developer Intent: Automatically ensure every processed workbook contains a theme by detecting absent themes and applying the built‑in default.
// Use Cases: Batch process multiple Excel files, adding a theme when none is present. | Integrate theme validation into an existing Aspose.Cells data‑processing pipeline. | Create a fallback theme for user‑generated workbooks that lack styling.
// AI Prompts: Generate C# code that iterates over Excel files, checks Workbook.Theme, and applies the default theme using Aspose.Cells. | Show how to use Workbook.CopyTheme to transfer a theme from a newly created workbook to an existing one, including error handling. | Explain the purpose of Workbook.Theme and how to guarantee a theme before performing further Excel manipulations with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookThemeProcessor
{
    // C# code that scans a list of Excel files, loads each workbook with Aspose.Cells, detects a missing Theme, creates a temporary workbook containing the built‑in default theme, copies that theme via Workbook.CopyTheme, and saves the updated file to an output folder.
    class Program
    {
        static void Main()
        {
            // List of workbook file paths to process
            List<string> workbookFiles = new List<string>
            {
                "Input1.xlsx",
                "Input2.xlsx",
                // Add more file paths as needed
            };

            // Ensure the output directory exists
            string outputDir = "Processed";
            Directory.CreateDirectory(outputDir);

            // Iterate through each workbook
            foreach (string filePath in workbookFiles)
            {
                try
                {
                    // Verify the input file exists
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Input file not found: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // If the workbook has no theme, copy the default theme from a new workbook
                    if (string.IsNullOrEmpty(workbook.Theme))
                    {
                        // Create a temporary workbook that contains the default theme
                        Workbook sourceWithTheme = new Workbook(FileFormatType.Xlsx);

                        // Copy the theme from the source workbook to the target workbook
                        workbook.CopyTheme(sourceWithTheme);
                    }

                    // Perform any additional processing here
                    // ...

                    // Save the workbook after ensuring it has a theme
                    string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
                    workbook.Save(outputPath);
                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("All workbooks processed.");
        }
    }
}
