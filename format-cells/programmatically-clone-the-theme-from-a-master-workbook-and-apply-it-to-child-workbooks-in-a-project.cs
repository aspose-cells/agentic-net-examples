// Title: Copy a master workbook theme to multiple child workbooks with Aspose.Cells for .NET (C#)
// Common Searches: how to use Aspose.Cells to copy an Excel theme from one workbook to another in C# | C# code for applying a master workbook theme to several child workbooks automatically | batch copy Excel theme across multiple files using Aspose.Cells .NET library | Aspose.Cells Workbook.CopyTheme method example for cloning themes | programmatically set the same theme for many XLSX files with Aspose.Cells
// Tags: Workbook.CopyTheme C# example | apply Excel theme Aspose.Cells | clone workbook theme .NET | batch theme copy across XLSX files | master theme propagation using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ThemeCopyExample
{
    // The sample loads a master Excel workbook, iterates through a list of child workbook paths, uses Workbook.CopyTheme to clone the master’s theme into each child workbook, saves the changes back to the original files, and includes file‑existence checks and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string masterPath = "MasterWorkbook.xlsx";

            // Verify master workbook exists
            if (!File.Exists(masterPath))
            {
                Console.WriteLine($"Master workbook not found: {masterPath}");
                return;
            }

            try
            {
                // Load the master workbook that contains the desired theme
                Workbook masterWorkbook = new Workbook(masterPath);

                // List of child workbooks to receive the cloned theme
                string[] childWorkbookPaths = new string[]
                {
                    "ChildWorkbook1.xlsx",
                    "ChildWorkbook2.xlsx",
                    "ChildWorkbook3.xlsx"
                };

                foreach (string childPath in childWorkbookPaths)
                {
                    try
                    {
                        // Verify child workbook exists
                        if (!File.Exists(childPath))
                        {
                            Console.WriteLine($"Child workbook not found: {childPath}");
                            continue;
                        }

                        // Load the child workbook
                        Workbook childWorkbook = new Workbook(childPath);

                        // Copy the theme from the master workbook to the child workbook
                        childWorkbook.CopyTheme(masterWorkbook);

                        // Save the child workbook, overwriting the original file
                        childWorkbook.Save(childPath, SaveFormat.Xlsx);

                        Console.WriteLine($"Theme applied and saved for: {childPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing '{childPath}': {ex.Message}");
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
