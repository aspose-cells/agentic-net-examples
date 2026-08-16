// Title: C# – Batch Protect VBA Projects in XLSM Files with Aspose.Cells
// Description: Scans a folder for *.xlsm workbooks, loads each file with Aspose.Cells, detects unprotected VBA projects and applies a password, then saves the workbook back in XLSM format. Ideal for automating macro security across multiple Excel files.
// Keywords: Aspose.Cells VBA protection | C# protect VBA project | batch password VBA macros | secure XLSM files programmatically | apply VBA password .NET | automate macro security | Excel VBA project encryption
// Common Searches: how to add a password to VBA projects in multiple xlsm files c# | batch protect unprotected VBA macros using Aspose.Cells | set VBA project password for all workbooks in a folder | check VBA project protection before applying password aspnet | automate VBA project encryption with Aspose.Cells
// Developer Intent: Automatically add a password to every unprotected VBA project in each XLSM workbook within a specified directory.
// Use Cases: Enforce macro security before distributing a library of Excel reports. | Meet compliance requirements by ensuring all macro‑enabled files have a VBA password. | Integrate into CI/CD pipelines to validate and protect VBA projects during build or release.
// AI Prompts: Write C# code that uses Aspose.Cells to protect VBA projects in all .xlsm files in a directory, skipping already protected projects. | Show how to extend the sample to also set a workbook opening password together with VBA protection. | Suggest advanced error handling, logging, and progress reporting for batch VBA project protection with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace ProtectVbaProjects
{
    // Scans a folder for *.xlsm workbooks, loads each file with Aspose.Cells, detects unprotected VBA projects and applies a password, then saves the workbook back in XLSM format. Ideal for automating macro security across multiple Excel files.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the XLSM files
            string folderPath = @"C:\Path\To\XlsmFolder";

            // Verify that the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Password to apply to unprotected VBA projects
            const string vbaPassword = "MyVbaPassword";

            // Iterate through all .xlsm files in the directory
            foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsm"))
            {
                // Ensure the file actually exists
                if (!File.Exists(filePath))
                {
                    continue;
                }

                try
                {
                    // Load the workbook from the file
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Access the VBA project associated with the workbook
                        VbaProject vbaProject = workbook.VbaProject;

                        // Protect the VBA project if it is not already protected
                        if (vbaProject != null && !vbaProject.IsProtected)
                        {
                            // Protect the VBA project without locking it for viewing
                            vbaProject.Protect(false, vbaPassword);

                            // Save the workbook, overwriting the original file
                            workbook.Save(filePath, SaveFormat.Xlsm);
                            Console.WriteLine($"Protected VBA project in: {Path.GetFileName(filePath)}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log any errors that occur while processing the file
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }
}
