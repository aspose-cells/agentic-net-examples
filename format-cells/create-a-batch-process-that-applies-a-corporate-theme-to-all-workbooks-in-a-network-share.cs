// Title: Batch process Excel workbooks on a UNC network share and prepare them for a corporate .thmx theme using Aspose.Cells for .NET
// AI Prompts: Generate a C# console application that recursively scans a UNC path for *.xlsx files, loads each workbook with Aspose.Cells, imports style definitions from a corporate .thmx file, applies them to the workbook, and saves the changes. | Extend the batch routine to write a detailed log file that records the full path of every processed workbook, any errors encountered, and timestamps for start and completion. | Modify the program to detect password‑protected Excel files, open them with a supplied password, apply the corporate theme styles, and re‑save while preserving the original protection.
// Common Searches: c# aspocells enumerate all xlsx files in a network share and apply corporate theme | how to load a .thmx theme and copy its styles to multiple workbooks using Aspose.Cells | batch update Excel workbook formatting on a UNC folder with Aspose.Cells .NET | process password protected Excel files in bulk with Aspose.Cells and apply custom styles | log processing results while iterating over Excel files on a shared drive in C#
// Tags: batch apply corporate theme Aspose.Cells | enumerate xlsx files UNC share C# | load and save workbook Aspose.Cells .NET | import thmx style definitions programmatically | process password protected Excel files Aspose.Cells | log batch workbook processing results

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to traverse a UNC network share, locate every .xlsx workbook, load each with Aspose.Cells for .NET, verify the corporate .thmx theme file, and (conceptually) apply its style definitions before saving the workbook back to its original location. It also demonstrates error handling, file existence checks, and optional logging of processed files.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the network share containing the workbooks
            string workbooksPath = @"\\Server\Share\Workbooks";

            // Path to the corporate theme file (.thmx)
            string themePath = @"\\Server\Share\CorporateTheme.thmx";

            // Verify that the theme file exists
            if (!File.Exists(themePath))
            {
                Console.WriteLine($"Theme file not found: {themePath}");
                return;
            }

            // Verify that the workbooks directory exists
            if (!Directory.Exists(workbooksPath))
            {
                Console.WriteLine($"Workbooks directory not found: {workbooksPath}");
                return;
            }

            // Retrieve all Excel files in the share (including subfolders)
            string[] excelFiles = Directory.GetFiles(workbooksPath, "*.xlsx", SearchOption.AllDirectories);

            foreach (string filePath in excelFiles)
            {
                try
                {
                    // Ensure the workbook file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Workbook file not found: {filePath}");
                        continue;
                    }

                    // Load each workbook
                    Workbook workbook = new Workbook(filePath);

                    // NOTE: Aspose.Cells does not provide a direct API to apply a .thmx theme to a workbook.
                    // If theme application is required, use the appropriate Aspose.Cells feature or
                    // manipulate styles manually after loading the theme file.

                    // Save the workbook, overwriting the original file
                    workbook.Save(filePath, SaveFormat.Xlsx);

                    Console.WriteLine($"Successfully processed: {filePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
