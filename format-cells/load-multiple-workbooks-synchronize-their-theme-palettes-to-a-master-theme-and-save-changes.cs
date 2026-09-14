// Title: Batch copy a master theme to multiple Excel .xlsx workbooks using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a master workbook and uses Aspose.Cells to transfer its theme to each workbook in a list, then saves the files. | Create a reusable C# method that takes a master workbook path and an IEnumerable of target workbook paths, applies the master theme with Aspose.Cells CopyTheme, and returns a processing summary. | Enhance the theme‑synchronization script to output a CSV log containing file name, success/failure status, and error details for each workbook.
// Common Searches: Aspose.Cells C# transfer theme from a source workbook to many target workbooks | How to apply the same Excel theme to several .xlsx files using .NET | Batch update Excel workbook themes programmatically with Aspose.Cells | C# script to synchronize Excel theme palettes across multiple workbooks | CopyTheme method example for processing multiple Excel files Aspose.Cells
// Tags: Aspose.Cells CopyTheme batch processing | C# apply master Excel theme | synchronize Excel theme palettes .NET | programmatic Excel theme copy Aspose.Cells | batch update workbook theme .xlsx

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads a designated master workbook, iterates over a list of other Excel files, copies the master workbook's theme to each using the Aspose.Cells CopyTheme method, saves the changes back to the original files, and handles missing files with console logging.
class ThemeSynchronizer
{
    static void Main()
    {
        // Paths to the workbooks to be synchronized
        List<string> workbookPaths = new List<string>
        {
            @"C:\Data\MasterWorkbook.xlsx",   // Master workbook containing the desired theme
            @"C:\Data\Workbook1.xlsx",
            @"C:\Data\Workbook2.xlsx",
            @"C:\Data\Workbook3.xlsx"
        };

        try
        {
            // Verify that the master workbook exists
            if (!File.Exists(workbookPaths[0]))
            {
                Console.WriteLine($"Master workbook not found: {workbookPaths[0]}");
                return;
            }

            // Load the master workbook (the first entry in the list)
            Workbook masterWorkbook = new Workbook(workbookPaths[0]);

            // Iterate over the remaining workbooks and apply the master theme
            for (int i = 1; i < workbookPaths.Count; i++)
            {
                string targetPath = workbookPaths[i];

                // Skip if the target file does not exist
                if (!File.Exists(targetPath))
                {
                    Console.WriteLine($"Target workbook not found, skipping: {targetPath}");
                    continue;
                }

                try
                {
                    // Load the target workbook
                    Workbook targetWorkbook = new Workbook(targetPath);

                    // Copy the theme from the master workbook to the target workbook
                    targetWorkbook.CopyTheme(masterWorkbook);

                    // Save the workbook (overwrites the original file)
                    targetWorkbook.Save(targetPath);
                    Console.WriteLine($"Theme synchronized for: {targetPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to synchronize theme for {targetPath}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
