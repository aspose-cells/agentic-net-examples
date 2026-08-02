// Title: C# – Detect Missing Workbook Theme and Apply a Default Theme with Aspose.Cells
// Description: A C# console example that loops through Excel files, loads each workbook with Aspose.Cells, checks the Workbook.Theme property, and when no theme is present creates a temporary workbook to obtain the built‑in theme. The theme is copied via Workbook.CopyTheme and the file is saved with a "_with_theme" suffix, enabling batch theme enforcement before further processing.
// Keywords: Aspose.Cells C# theme detection | apply default Excel theme programmatically | Workbook.Theme check | CopyTheme example | batch apply Excel theme .NET | Aspose.Cells workbook theme | C# Excel theme automation | default theme workbook copy
// Common Searches: how to check if an Excel workbook has a theme using Aspose.Cells | apply a default theme to workbooks lacking one in C# | copy theme from a template workbook to another workbook Aspose.Cells | Aspose.Cells detect missing theme and set default | C# batch add theme to multiple Excel files
// Developer Intent: Identify Excel workbooks without a theme and programmatically assign the default theme before saving.
// Use Cases: Ensure all corporate reports share a unified visual style by batch‑applying a standard theme to incoming spreadsheets. | Pre‑process user‑uploaded Excel files in a web service, enforcing a default theme to guarantee consistent rendering. | Automate theme compliance for legacy workbooks before performing data extraction or analytics.
// AI Prompts: Generate a C# method that receives a file path, verifies Workbook.Theme, and applies the default theme with Aspose.Cells if missing. | Show how to copy a custom theme from a template workbook to a list of target workbooks, including robust error handling. | Explain how to confirm that a theme was successfully applied after using Workbook.CopyTheme in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookThemeProcessor
{
    // A C# console example that loops through Excel files, loads each workbook with Aspose.Cells, checks the Workbook.Theme property, and when no theme is present creates a temporary workbook to obtain the built‑in theme. The theme is copied via Workbook.CopyTheme and the file is saved with a "_with_theme" suffix, enabling batch theme enforcement before further processing.
    class Program
    {
        static void Main(string[] args)
        {
            // List of workbook file paths to process
            string[] workbookFiles = new string[]
            {
                "input1.xlsx",
                "input2.xlsx"
                // Add more file paths as needed
            };

            foreach (string filePath in workbookFiles)
            {
                // Verify that the input file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File '{filePath}' not found. Skipping.");
                    continue;
                }

                try
                {
                    // Load the workbook from the file
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Check if the workbook has a theme assigned
                        if (string.IsNullOrEmpty(workbook.Theme))
                        {
                            // Create a temporary workbook that contains the default theme
                            using (Workbook defaultThemeWorkbook = new Workbook())
                            {
                                // Copy the default theme into the target workbook
                                workbook.CopyTheme(defaultThemeWorkbook);
                            }

                            Console.WriteLine($"Default theme applied to '{Path.GetFileName(filePath)}'.");
                        }
                        else
                        {
                            Console.WriteLine($"'{Path.GetFileName(filePath)}' already has theme: {workbook.Theme}");
                        }

                        // Save the workbook (creates a new file with a suffix)
                        string outputPath = Path.Combine(
                            Path.GetDirectoryName(filePath) ?? string.Empty,
                            Path.GetFileNameWithoutExtension(filePath) + "_with_theme.xlsx");

                        workbook.Save(outputPath, SaveFormat.Xlsx);
                        Console.WriteLine($"Workbook saved as '{Path.GetFileName(outputPath)}'.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }
        }
    }
}
