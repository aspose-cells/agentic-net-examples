// Title: Copy a master theme to multiple Excel workbooks with Aspose.Cells for .NET
// Common Searches: how to copy a theme from a master Excel file to other workbooks using Aspose.Cells C# | batch apply a shared theme to several .xlsx files in a .NET application | Aspose.Cells example for synchronizing workbook themes across multiple spreadsheets | C# code to propagate Excel theme changes to multiple workbooks automatically | using Workbook.CopyTheme to update themes in a list of Excel files
// Tags: Workbook.CopyTheme method | master theme propagation Excel .NET | batch update Excel theme Aspose.Cells | copy theme between workbooks C# | synchronize workbook palettes Aspose.Cells | automate Excel theme changes C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads a master workbook, iterates through a collection of target .xlsx files, copies the master workbook's theme to each using Workbook.CopyTheme, and saves the updated workbooks, enabling batch theme synchronization across multiple Excel files.
class ThemeSynchronizer
{
    static void Main()
    {
        // Path to the workbook that contains the master theme
        string masterWorkbookPath = "MasterTheme.xlsx";

        // List of workbook files whose themes need to be synchronized with the master theme
        List<string> targetWorkbookPaths = new List<string>
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx"
        };

        // Verify master workbook exists
        if (!File.Exists(masterWorkbookPath))
        {
            Console.WriteLine($"Master workbook not found: {masterWorkbookPath}");
            return;
        }

        Workbook masterWorkbook;
        try
        {
            // Load the master workbook
            masterWorkbook = new Workbook(masterWorkbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load master workbook: {ex.Message}");
            return;
        }

        // Iterate through each target workbook, apply the master theme, and save the changes
        foreach (string targetPath in targetWorkbookPaths)
        {
            // Verify target workbook exists
            if (!File.Exists(targetPath))
            {
                Console.WriteLine($"Target workbook not found: {targetPath}");
                continue;
            }

            try
            {
                // Load the target workbook
                Workbook targetWorkbook = new Workbook(targetPath);

                // Copy the theme from the master workbook to the target workbook
                targetWorkbook.CopyTheme(masterWorkbook);

                // Save the workbook, overwriting the original file
                targetWorkbook.Save(targetPath, SaveFormat.Xlsx);
                Console.WriteLine($"Theme synchronized for: {targetPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {targetPath}: {ex.Message}");
            }
        }
    }
}
