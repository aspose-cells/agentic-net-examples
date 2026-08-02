// Title: Copy Excel theme from a master workbook to multiple workbooks with Aspose.Cells for .NET
// Description: C# example that loads a master Excel file, iterates over a list of child workbooks, and uses Aspose.Cells Workbook.CopyTheme to clone the master theme into each child workbook. Includes file‑existence checks and exception handling for robust bulk processing.
// Keywords: Aspose.Cells CopyTheme | C# Excel theme cloning | apply master workbook theme | bulk Excel theme update | .NET Excel theme automation | copy Excel theme programmatically | Workbook.CopyTheme example
// Common Searches: Aspose.Cells copy theme between workbooks | How to apply a master Excel theme to many files in C# | Bulk theme transfer using Aspose.Cells | CopyTheme method C# example | Programmatically clone Excel theme
// Developer Intent: Programmatically duplicate a master workbook's theme and apply it to a collection of child workbooks.
// Use Cases: Enforce corporate branding across all departmental reports. | Synchronize visual styles after a design refresh in an automated reporting pipeline. | Prepare a batch of generated Excel files with a consistent theme without manual editing.
// AI Prompts: Generate C# code that loads a master Excel file and copies its theme to a list of target workbooks using Aspose.Cells, with error handling for missing or corrupt files. | Explain how Workbook.CopyTheme works and what style elements are transferred when cloning a theme. | Show how to log the success or failure of theme application for each workbook in a bulk operation.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// C# example that loads a master Excel file, iterates over a list of child workbooks, and uses Aspose.Cells Workbook.CopyTheme to clone the master theme into each child workbook. Includes file‑existence checks and exception handling for robust bulk processing.
public class ThemeCloner
{
    // Copies the theme from a master workbook to each child workbook in the list.
    public static void ApplyThemeToChildren(string masterPath, IEnumerable<string> childPaths)
    {
        if (!File.Exists(masterPath))
        {
            Console.WriteLine($"Master workbook not found: {masterPath}");
            return;
        }

        Workbook masterWorkbook;
        try
        {
            masterWorkbook = new Workbook(masterPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load master workbook: {ex.Message}");
            return;
        }

        foreach (var childPath in childPaths)
        {
            if (!File.Exists(childPath))
            {
                Console.WriteLine($"Child workbook not found, skipping: {childPath}");
                continue;
            }

            Workbook childWorkbook;
            try
            {
                childWorkbook = new Workbook(childPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load child workbook '{childPath}': {ex.Message}");
                continue;
            }

            try
            {
                // Apply the master theme to the child workbook.
                childWorkbook.CopyTheme(masterWorkbook);
                // Save the child workbook, overwriting the original file.
                childWorkbook.Save(childPath);
                Console.WriteLine($"Theme applied and saved to: {childPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{childPath}': {ex.Message}");
            }
        }
    }

    // Example entry point demonstrating usage.
    public static void Main()
    {
        // Path to the workbook that holds the source theme.
        string masterFile = "MasterTemplate.xlsx";

        // List of child workbooks that should receive the theme.
        List<string> childFiles = new List<string>
        {
            "Child1.xlsx",
            "Child2.xlsx",
            "Child3.xlsx"
        };

        // Apply the theme to all child workbooks.
        ApplyThemeToChildren(masterFile, childFiles);
    }
}
