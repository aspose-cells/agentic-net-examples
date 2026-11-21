// Title: Protect unprotected VBA projects in multiple XLSM files with Aspose.Cells for .NET
// Common Searches: how to add a password to VBA projects in all .xlsm files using Aspose.Cells C# | batch protect VBA macros in a folder of Excel macro-enabled workbooks with .NET | C# code to protect only unprotected VBA projects in XLSM workbooks | iterate through a directory and secure VBA projects using Aspose.Cells | apply VBA project password programmatically to multiple workbooks
// Tags: protect VBA project Aspose.Cells | batch VBA project password C# | process XLSM files directory Aspose | VbaProject.Protect method example | secure macro-enabled workbooks programmatically | apply password to VBA project .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject

// The sample scans a given folder for *.xlsm files, loads each workbook with Aspose.Cells, checks for an existing but unprotected VBA project, applies a password using VbaProject.Protect(true, password), and saves the workbook back in XLSM format while handling missing files and exceptions.
class ProtectVbaProjects
{
    static void Main()
    {
        // Directory containing the XLSM files
        string folderPath = @"C:\Path\To\Your\XlsmFolder";

        // Verify that the directory exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Password to apply to VBA projects
        string vbaPassword = "YourVbaPassword";

        // Get all .xlsm files in the directory
        string[] files = Directory.GetFiles(folderPath, "*.xlsm", SearchOption.TopDirectoryOnly);

        foreach (string filePath in files)
        {
            // Ensure the file still exists before processing
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found (skipped): {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Access the VBA project (if any)
                VbaProject vbaProject = workbook.VbaProject;

                // Proceed only if a VBA project exists and is not already protected
                if (vbaProject != null && !vbaProject.IsProtected)
                {
                    // Protect the VBA project with the specified password
                    // The first argument indicates that a password is required
                    vbaProject.Protect(true, vbaPassword);

                    // Save the workbook back to the same file (preserving XLSM format)
                    workbook.Save(filePath, SaveFormat.Xlsm);
                    Console.WriteLine($"Protected VBA project in: {filePath}");
                }
                else
                {
                    Console.WriteLine($"No unprotected VBA project found in: {filePath}");
                }
            }
            catch (Exception ex)
            {
                // Log any errors but continue processing other files
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }
    }
}
