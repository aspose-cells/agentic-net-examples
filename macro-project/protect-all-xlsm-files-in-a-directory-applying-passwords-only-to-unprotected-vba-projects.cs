// Title: Protect VBA Projects in All XLSM Files in a Folder – Aspose.Cells C# Example
// Description: A C# script that scans a specified folder, loads each *.xlsm workbook with Aspose.Cells, checks if the VBA project is unprotected, applies a password, and saves the file back in place.
// Keywords: Aspose.Cells | C# | .NET | VBA project protection | XLSM batch processing | Excel macro password | directory file iteration | Excel automation | macro security | protect unprotected VBA
// Common Searches: batch protect VBA projects in XLSM files C# | Aspose.Cells protect VBA macro password | apply password to all VBA projects in a folder | C# script to secure macro‑enabled Excel workbooks | iterate XLSM files and protect VBA using Aspose
// Developer Intent: Loop through every XLSM file in a given directory, detect VBA projects without a password, protect them with a specified password, and overwrite the original files.
// Use Cases: Enforce macro security across corporate Excel templates before distribution. | Integrate into CI/CD pipelines to guarantee all generated reports have protected VBA code. | Run a maintenance job on shared drives that adds passwords to any unprotected VBA projects.
// AI Prompts: Write C# code using Aspose.Cells that protects only unprotected VBA projects in all XLSM files of a folder. | Show how to extend the script to process subfolders recursively while keeping the same protection logic. | Suggest a logging strategy that records files already protected versus those newly password‑protected.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// A C# script that scans a specified folder, loads each *.xlsm workbook with Aspose.Cells, checks if the VBA project is unprotected, applies a password, and saves the file back in place.
class ProtectVbaProjectsInDirectory
{
    static void Main()
    {
        try
        {
            // Directory containing the XLSM files
            string folderPath = @"C:\Path\To\XlsmFolder";

            // Verify that the directory exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Directory not found: {folderPath}");
                return;
            }

            // Password to apply to unprotected VBA projects
            const string vbaPassword = "MyVbaPassword";

            // Get all .xlsm files in the directory (non‑recursive)
            string[] xlsmFiles = Directory.GetFiles(folderPath, "*.xlsm", SearchOption.TopDirectoryOnly);

            foreach (string filePath in xlsmFiles)
            {
                // Ensure the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(filePath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // If the VBA project is not protected, protect it with the specified password
                if (!vbaProject.IsProtected)
                {
                    // Protect without locking for viewing (islockedForViewing = false)
                    vbaProject.Protect(false, vbaPassword);
                }

                // Save the workbook back to the same file (lifecycle rule: save)
                workbook.Save(filePath, SaveFormat.Xlsm);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
