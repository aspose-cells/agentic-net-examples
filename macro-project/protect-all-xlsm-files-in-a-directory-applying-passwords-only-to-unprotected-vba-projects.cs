// Title: Protect unprotected VBA projects in multiple XLSM workbooks using Aspose.Cells for .NET
// Common Searches: C# code to add a password to VBA projects in all .xlsm files in a folder with Aspose.Cells | How to batch protect macro-enabled Excel workbooks VBA projects using Aspose.Cells .NET | Iterate through a directory of XLSM files and set VBA project password programmatically | Aspose.Cells protect VBA project only when it is not already protected | Apply password to VBA project in Excel macro-enabled files using C#
// Tags: batch protect VBA project Aspose.Cells | set VBA password XLSM C# | VbaProject.Protect method Aspose.Cells | iterate XLSM files directory C# | check VbaProject.IsProtected Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The sample scans a specified folder for *.xlsm files, loads each workbook with Aspose.Cells, checks if a VBA project exists and is unprotected, applies a password using VbaProject.Protect, and saves the workbook back in XLSM format, logging the result for each file.
class ProtectVbaProjects
{
    static void Main()
    {
        // Directory containing the XLSM files
        string folderPath = @"C:\Path\To\Folder";

        // Verify the directory exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Password to apply to unprotected VBA projects
        const string vbaPassword = "YourVbaPassword";

        // Get all XLSM files in the directory
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
                // Load the workbook (VBA project is loaded automatically for XLSM files)
                Workbook workbook = new Workbook(filePath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // If a VBA project exists and is not already protected, protect it
                if (vbaProject != null && !vbaProject.IsProtected)
                {
                    // Protect the VBA project with the specified password
                    vbaProject.Protect(true, vbaPassword);

                    // Save the workbook, preserving the VBA project
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
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
