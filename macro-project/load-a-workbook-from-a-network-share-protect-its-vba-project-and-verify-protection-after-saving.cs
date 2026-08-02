// Title: Protect VBA Project in a Macro‑Enabled Workbook Loaded from a Network Share (C# Aspose.Cells)
// Description: Loads an .xlsm file from a UNC path (or creates one), locks its VBA project with a password, saves the workbook, reloads it, and confirms protection using Aspose.Cells VbaProject APIs.
// Keywords: Aspose.Cells C# | VBA project protection | macro-enabled workbook | UNC network share | XLSM password lock | ValidatePassword API | programmatic VBA lock | load workbook from network | save protected workbook
// Common Searches: Aspose.Cells protect VBA project C# | load .xlsm from UNC path Aspose.Cells | set VBA password programmatically .NET | verify VBA protection after saving | create macro workbook on network share
// Developer Intent: Open a macro‑enabled workbook from a network location, apply password protection to its VBA project, save the file, and ensure the protection persists.
// Use Cases: Automate security of VBA code in shared reports before distribution. | Enforce compliance by programmatically locking macro workbooks stored on a file server. | Audit saved workbooks to confirm VBA protection and password validity.
// AI Prompts: Write C# code with Aspose.Cells to open an .xlsm file from a UNC path, protect its VBA project with a password, and save it. | Show how to reload a saved workbook and check the IsProtected flag and password validation for the VBA project. | Explain handling of workbooks that lack a VBA project when applying protection using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionDemo
{
    // Loads an .xlsm file from a UNC path (or creates one), locks its VBA project with a password, saves the workbook, reloads it, and confirms protection using Aspose.Cells VbaProject APIs.
    class Program
    {
        static void Main()
        {
            // Path to the workbook on a network share (replace with actual path)
            string networkPath = @"\\Server\Share\SampleWorkbook.xlsm";

            // Output path for the protected workbook
            string outputPath = @"C:\Temp\ProtectedWorkbook.xlsm";

            try
            {
                // Ensure the directory for the network path exists
                string networkDir = Path.GetDirectoryName(networkPath);
                if (!string.IsNullOrEmpty(networkDir) && !Directory.Exists(networkDir))
                {
                    Directory.CreateDirectory(networkDir);
                }

                // Ensure the directory for the output path exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                Workbook workbook;

                // Load workbook from network share if it exists; otherwise create a new macro-enabled workbook
                if (File.Exists(networkPath))
                {
                    workbook = new Workbook(networkPath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Save(networkPath, SaveFormat.Xlsm);
                    workbook = new Workbook(networkPath);
                }

                // Protect the VBA project (lock for viewing = true) with a password
                workbook.VbaProject?.Protect(true, "vbaPassword123");

                // Save the workbook (macro-enabled format required for VBA)
                workbook.Save(outputPath, SaveFormat.Xlsm);

                // Reload the saved workbook to verify protection status
                Workbook reloadedWorkbook = new Workbook(outputPath);
                VbaProject vbaProject = reloadedWorkbook.VbaProject;

                // Output protection information
                Console.WriteLine("VBA Project IsProtected: " + vbaProject?.IsProtected);

                // Validate the password (if a VBA project exists)
                if (vbaProject != null)
                {
                    bool isPasswordValid = vbaProject.ValidatePassword("vbaPassword123");
                    Console.WriteLine("Password validation result: " + isPasswordValid);
                }
                else
                {
                    Console.WriteLine("No VBA project found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
