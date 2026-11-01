// Title: Programmatically copy a master workbook’s theme to multiple Excel files using Aspose.Cells for .NET
// Common Searches: C# code to copy an Excel workbook theme to other workbooks with Aspose.Cells | How to batch apply a master theme to several .xlsx files in .NET | Aspose.Cells example for synchronizing theme palettes across multiple workbooks | Automate Excel theme transfer for multiple files using C# and Aspose | Copy theme from one workbook to many using Aspose.Cells CopyTheme method
// Tags: Aspose.Cells Workbook.CopyTheme usage | batch Excel theme synchronization .NET | copy theme between Excel workbooks C# | automate Excel theme palette transfer | apply master theme to multiple .xlsx files

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a master workbook, iterates through a list of target Excel files, copies the master workbook's theme to each using Workbook.CopyTheme, and saves the updated files back in XLSX format.
class ThemeSynchronizer
{
    static void Main()
    {
        try
        {
            // Path to the master workbook whose theme will be used
            string masterWorkbookPath = "Master.xlsx";

            // Verify master workbook exists
            if (!File.Exists(masterWorkbookPath))
            {
                Console.WriteLine($"Master workbook not found: {masterWorkbookPath}");
                return;
            }

            // List of workbook file paths that need to be synchronized with the master theme
            List<string> workbookPaths = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
                // Add more paths as needed
            };

            // Load the master workbook
            Workbook masterWorkbook = new Workbook(masterWorkbookPath);

            // Loop through each target workbook, apply the master theme, and save the changes
            foreach (string path in workbookPaths)
            {
                // Verify target workbook exists
                if (!File.Exists(path))
                {
                    Console.WriteLine($"Target workbook not found: {path}");
                    continue;
                }

                // Load the target workbook
                Workbook targetWorkbook = new Workbook(path);

                // Apply the master theme to the target workbook
                targetWorkbook.CopyTheme(masterWorkbook);

                // Save the workbook, overwriting the original file
                targetWorkbook.Save(path, SaveFormat.Xlsx);
                Console.WriteLine($"Theme synchronized for: {path}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
