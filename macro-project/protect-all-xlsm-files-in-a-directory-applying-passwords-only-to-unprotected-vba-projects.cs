// Title: Batch protect unprotected VBA projects in multiple XLSM workbooks using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that scans a directory, loads each .xlsm workbook, checks for a VBA project, and applies a password only when the project is not already protected. | Write a reusable method `ProtectVbaInFolder(string folderPath, string vbaPassword)` that iterates over all XLSM files, protects any unprotected VBA projects, and saves the workbooks preserving macros. | Enhance the batch script with detailed console logging that reports each file name, whether its VBA project was protected or already secured, and any errors encountered.
// Common Searches: aspnet c# batch protect VBA projects in all xlsm files in a directory using Aspose.Cells | how to add password to unprotected macro projects in multiple Excel workbooks with Aspose.Cells | C# loop through .xlsm files and set VBA project password only if not already set | protect VBA project programmatically in Excel files without affecting existing protection Aspose.Cells
// Tags: Aspose.Cells protect VBA project | batch VBA password protection XLSM | C# iterate XLSM files Aspose.Cells | VbaProject.IsProtected check | save workbook as Xlsm after VBA protection

using System;
using System.IO;
using Aspose.Cells;

// C# program that scans a specified folder for .xlsm files, loads each workbook with Aspose.Cells, detects a VBA project, applies a password only if the project is not already protected, and saves the workbook back in XLSM format while preserving existing macros.
class ProtectVbaProjects
{
    static void Main()
    {
        try
        {
            // Directory containing the XLSM files
            string folderPath = @"C:\Path\To\Folder";

            // Verify that the directory exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Password to apply to unprotected VBA projects
            string vbaPassword = "YourVbaPassword";

            // Iterate through all .xlsm files in the directory
            foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsm"))
            {
                // Ensure the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                // Load the workbook (preserves existing content and macros)
                Workbook workbook = new Workbook(filePath);

                // Ensure the workbook contains a VBA project
                if (workbook.VbaProject != null)
                {
                    // Check if the VBA project is already protected
                    if (!workbook.VbaProject.IsProtected)
                    {
                        // Protect the VBA project with the specified password
                        // The first argument indicates that a password is required
                        workbook.VbaProject.Protect(true, vbaPassword);

                        // Save the workbook back to the same file, preserving the XLSM format
                        workbook.Save(filePath, SaveFormat.Xlsm);
                        Console.WriteLine($"Protected VBA project in: {filePath}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
