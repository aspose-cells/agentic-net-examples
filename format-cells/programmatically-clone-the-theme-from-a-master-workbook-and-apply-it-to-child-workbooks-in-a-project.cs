// Title: Programmatically copy a master workbook theme to multiple Excel files using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# routine that opens a master Excel file and uses Aspose.Cells' Workbook.CopyTheme to transfer its theme to each workbook in a supplied list, then saves the changes. | Create a function that takes a source workbook path and an array of destination paths, clones the source theme onto every destination workbook using Aspose.Cells, and returns any processing errors. | Develop a .NET script that iterates over multiple .xlsx files, applying the theme from a reference workbook via Aspose.Cells without manual intervention.
// Common Searches: Aspose.Cells copy theme from master workbook to other workbooks C# | How to programmatically apply the same Excel theme to many files using .NET | Batch copy Excel workbook theme with Aspose.Cells CopyTheme method | C# automate theme propagation across multiple .xlsx documents | Apply master workbook theme to child workbooks in a project using Aspose.Cells
// Tags: Aspose.Cells Workbook.CopyTheme example | copy Excel theme programmatically C# | batch apply workbook theme Aspose.Cells | clone master workbook theme .NET | automate Excel theme propagation Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads a master Excel workbook, iterates through a list of child workbook paths, copies the master’s theme to each child using Workbook.CopyTheme, saves the updated files, and logs missing files or errors.
class ThemeCloner
{
    static void Main(string[] args)
    {
        // Path to the master workbook that contains the desired theme
        string masterPath = @"C:\Project\MasterWorkbook.xlsx";

        // Verify master workbook exists
        if (!File.Exists(masterPath))
        {
            Console.WriteLine($"Master workbook not found: {masterPath}");
            return;
        }

        try
        {
            // Load the master workbook
            using (Workbook masterWorkbook = new Workbook(masterPath))
            {
                // List of child workbook file paths that need to receive the master theme
                List<string> childWorkbookPaths = new List<string>
                {
                    @"C:\Project\Child1.xlsx",
                    @"C:\Project\Child2.xlsx",
                    @"C:\Project\Child3.xlsx"
                    // Add more paths as needed
                };

                // Iterate through each child workbook, copy the theme from the master, and save
                foreach (string childPath in childWorkbookPaths)
                {
                    // Verify child workbook exists
                    if (!File.Exists(childPath))
                    {
                        Console.WriteLine($"Child workbook not found, skipping: {childPath}");
                        continue;
                    }

                    try
                    {
                        // Load the child workbook
                        using (Workbook childWorkbook = new Workbook(childPath))
                        {
                            // Copy the theme from the master workbook to the child workbook
                            childWorkbook.CopyTheme(masterWorkbook);

                            // Save the updated child workbook (overwrites the original file)
                            childWorkbook.Save(childPath);
                            Console.WriteLine($"Theme applied and saved: {childPath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing child workbook '{childPath}': {ex.Message}");
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
