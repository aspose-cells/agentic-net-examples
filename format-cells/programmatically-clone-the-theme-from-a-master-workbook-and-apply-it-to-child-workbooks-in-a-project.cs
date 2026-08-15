// Title: Copy Excel Theme from a Master Workbook to Multiple Workbooks using Aspose.Cells for .NET
// Description: A C# example that loads a master Excel file, reads its theme, and uses Aspose.Cells Workbook.CopyTheme to apply the same branding to a list of child workbooks, saving each result with a distinct filename.
// Keywords: Aspose.Cells | CopyTheme | C# Excel theme | master workbook theme | apply theme to multiple workbooks | batch theme copy | Excel branding automation | Aspose.Cells .NET example | theme cloning | Workbook.CopyTheme
// Common Searches: Aspose.Cells copy theme C# | How to apply master Excel theme to other files using Aspose | Workbook.CopyTheme method example | Batch copy Excel theme with Aspose.Cells | C# program to clone workbook theme | Set corporate theme for multiple spreadsheets .NET
// Developer Intent: Transfer the theme from a single master Excel file to several child workbooks programmatically.
// Use Cases: Enforce corporate branding across all generated reports. | Refresh the visual style of legacy spreadsheets after a brand update. | Create new workbooks that inherit a standard theme while keeping existing data intact. | Automate theme consistency in a CI/CD pipeline for Excel report generation.
// AI Prompts: Generate a reusable C# function that takes a master workbook path and a list of target workbook paths, copies the master theme using Aspose.Cells, and saves each file with a configurable suffix. | Detail which theme components (colors, fonts, effects, cell styles) are copied by Workbook.CopyTheme in Aspose.Cells. | Provide a pattern for logging, retry, and exception handling when processing hundreds of workbooks with theme copying.

using System;
using System.IO;
using Aspose.Cells;

namespace ThemeCloner
{
    // A C# example that loads a master Excel file, reads its theme, and uses Aspose.Cells Workbook.CopyTheme to apply the same branding to a list of child workbooks, saving each result with a distinct filename.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the master workbook that contains the desired theme
            string masterPath = "MasterWorkbook.xlsx";

            // Verify that the master workbook exists
            if (!File.Exists(masterPath))
            {
                Console.WriteLine($"Master workbook not found: '{masterPath}'.");
                return;
            }

            try
            {
                // Load the master workbook (theme will be read from this file)
                using (Workbook masterWorkbook = new Workbook(masterPath))
                {
                    // Define the list of child workbook file paths that need the master theme applied
                    string[] childWorkbookPaths = new string[]
                    {
                        "ChildWorkbook1.xlsx",
                        "ChildWorkbook2.xlsx",
                        "ChildWorkbook3.xlsx"
                    };

                    foreach (string childPath in childWorkbookPaths)
                    {
                        // Verify that the child workbook exists
                        if (!File.Exists(childPath))
                        {
                            Console.WriteLine($"Child workbook not found: '{childPath}'. Skipping.");
                            continue;
                        }

                        try
                        {
                            // Load the child workbook
                            using (Workbook childWorkbook = new Workbook(childPath))
                            {
                                // Copy the theme from the master workbook to the child workbook
                                childWorkbook.CopyTheme(masterWorkbook);

                                // Determine output path (save to a new file)
                                string outputDir = Path.GetDirectoryName(childPath) ?? Directory.GetCurrentDirectory();
                                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(childPath);
                                string newFileName = $"{fileNameWithoutExt}_WithMasterTheme.xlsx";
                                string fullOutputPath = Path.Combine(outputDir, newFileName);

                                // Ensure the output directory exists
                                if (!Directory.Exists(outputDir))
                                {
                                    Directory.CreateDirectory(outputDir);
                                }

                                // Save the child workbook with the applied theme
                                childWorkbook.Save(fullOutputPath);
                                Console.WriteLine($"Theme copied to '{fullOutputPath}'.");
                            }
                        }
                        catch (Exception exChild)
                        {
                            Console.WriteLine($"Error processing '{childPath}': {exChild.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading master workbook: {ex.Message}");
            }
        }
    }
}
