// Title: Batch Synchronize Excel Workbook Themes with a Master Theme Using Aspose.Cells for .NET (C#)
// Description: C# example that loads a master workbook, iterates over multiple target workbooks, copies the master theme to each using Workbook.CopyTheme, saves the changes in‑place, and gracefully handles missing files and runtime errors.
// Keywords: Aspose.Cells CopyTheme | C# Excel theme synchronization | apply master theme to workbooks | bulk Excel theme update .NET | Excel workbook theme copy programmatically | standardize Excel branding | theme palette synchronization | Aspose.Cells batch processing
// Common Searches: Aspose.Cells copy theme from one workbook to another C# | Batch update Excel theme for multiple files using Aspose.Cells | How to apply a master Excel theme to several workbooks programmatically | C# example for synchronizing workbook themes with Aspose.Cells | Automate Excel theme consistency across reports .NET
// Developer Intent: Copy a master workbook's theme to a list of target workbooks and overwrite the originals.
// Use Cases: Enforce corporate branding by applying a single color palette to all generated reports. | Prepare a suite of template files with a consistent theme before distribution to users. | Normalize themes of spreadsheets received from external partners to maintain visual uniformity.
// AI Prompts: Generate C# code that loads a master Excel file and applies its theme to an array of workbook paths using Aspose.Cells, including detailed error handling and logging. | Suggest enhancements to the theme synchronization script to record a processing summary (files updated, skipped, errors) and to support custom output directories.

using System;
using System.IO;
using Aspose.Cells;

namespace ThemeSynchronizationDemo
{
    // C# example that loads a master workbook, iterates over multiple target workbooks, copies the master theme to each using Workbook.CopyTheme, saves the changes in‑place, and gracefully handles missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the master workbook whose theme will be applied to others
                string masterWorkbookPath = "master.xlsx";

                // Verify that the master workbook exists
                if (!File.Exists(masterWorkbookPath))
                {
                    Console.WriteLine($"Error: Master workbook not found at '{masterWorkbookPath}'.");
                    return;
                }

                // Paths of workbooks that need to be synchronized with the master theme
                string[] targetWorkbookPaths = new string[]
                {
                    "workbook1.xlsx",
                    "workbook2.xlsx",
                    "workbook3.xlsx"
                };

                // Load the master workbook (theme is automatically loaded with the file)
                Workbook masterWorkbook = new Workbook(masterWorkbookPath);

                // Iterate through each target workbook, copy the master theme, and save the changes
                foreach (string targetPath in targetWorkbookPaths)
                {
                    // Skip if the target workbook does not exist
                    if (!File.Exists(targetPath))
                    {
                        Console.WriteLine($"Warning: Target workbook not found at '{targetPath}'. Skipping.");
                        continue;
                    }

                    // Load the target workbook
                    Workbook targetWorkbook = new Workbook(targetPath);

                    // Copy the theme from the master workbook to the target workbook
                    targetWorkbook.CopyTheme(masterWorkbook);

                    // Save the updated workbook (overwrites the original file)
                    targetWorkbook.Save(targetPath);
                }

                Console.WriteLine("Theme synchronization completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
