// Title: Programmatically copy a master workbook theme to multiple child workbooks using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy theme from master workbook to multiple workbooks | How to programmatically apply a workbook theme to several Excel files using .NET | Batch apply Excel theme with Aspose.Cells CopyTheme method | Clone Excel theme across workbooks in C# with Aspose.Cells | Automate theme synchronization for multiple spreadsheets using Aspose.Cells
// Tags: CopyTheme method Aspose.Cells | apply workbook theme .NET | batch theme update Excel Aspose.Cells | master workbook theme propagation C# | clone Excel theme across files Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads a master Excel workbook, iterates through a list of child workbook paths, loads each existing child workbook, copies the master workbook's theme using the CopyTheme method, and saves the updated child workbook back to its original file.
class Program
{
    static void Main()
    {
        const string masterPath = "Master.xlsx";

        // Verify that the master workbook exists before attempting to load it.
        if (!File.Exists(masterPath))
        {
            Console.WriteLine($"Error: Master workbook '{masterPath}' not found.");
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

        // List of child workbooks that should receive the master theme.
        List<string> childWorkbookPaths = new List<string>
        {
            "Child1.xlsx",
            "Child2.xlsx",
            "Child3.xlsx"
        };

        foreach (string childPath in childWorkbookPaths)
        {
            // Skip processing if the child file does not exist.
            if (!File.Exists(childPath))
            {
                Console.WriteLine($"Warning: Child workbook '{childPath}' not found. Skipping.");
                continue;
            }

            try
            {
                // Load the child workbook.
                Workbook childWorkbook = new Workbook(childPath);

                // Apply the theme from the master workbook.
                childWorkbook.CopyTheme(masterWorkbook);

                // Overwrite the original file with the updated workbook.
                childWorkbook.Save(childPath, SaveFormat.Xlsx);

                Console.WriteLine($"Theme applied and saved for '{childPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{childPath}': {ex.Message}");
            }
        }
    }
}
