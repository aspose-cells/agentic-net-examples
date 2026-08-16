// Title: Password‑protect a VBA project in an XLSM workbook using Aspose.Cells for .NET
// Description: Load an .xlsm file with Aspose.Cells, verify the VBA project's IsProtected flag, apply Workbook.VbaProject.Protect with a strong password only when needed, and save the workbook as a protected macro‑enabled file.
// Keywords: Aspose.Cells C# | protect VBA project | XLSM password protection | Workbook.VbaProject.Protect | macro‑enabled workbook security | C# Excel automation | .NET Excel library
// Common Searches: How to add a password to a VBA project with Aspose.Cells | C# code to protect unprotected VBA macros in XLSM | Aspose.Cells check VBA IsProtected before saving | Secure macro‑enabled Excel files using .NET | Apply password to VBA project programmatically
// Developer Intent: Secure a VBA project with a strong password only if it is currently unprotected.
// Use Cases: Automated processing of macro‑enabled workbooks that must be locked before distribution. | CI/CD pipelines that enforce VBA protection on generated XLSM reports. | Pre‑flight validation of Excel files to ensure macro security before archival.
// AI Prompts: Generate C# code that loads an XLSM file, checks Workbook.VbaProject.IsProtected, and calls Protect(false, password) when false. | Create error‑handling logic for missing input files and protection failures in Aspose.Cells VBA protection scripts. | Explain the meaning of the isLockedForViewing parameter in Workbook.VbaProject.Protect and when to set it to true or false.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsVbaProtection
{
    // Load an .xlsm file with Aspose.Cells, verify the VBA project's IsProtected flag, apply Workbook.VbaProject.Protect with a strong password only when needed, and save the workbook as a protected macro‑enabled file.
    public class ProtectVbaProject
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsm";
            const string outputPath = "output_protected.xlsm";
            const string password = "Str0ngP@ssw0rd!2026";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook that contains a VBA project
                Workbook workbook = new Workbook(inputPath);

                // Protect the VBA project if it is not already protected
                if (!workbook.VbaProject.IsProtected)
                {
                    // isLockedForViewing = false (project can be opened, but editing is restricted)
                    workbook.VbaProject.Protect(false, password);
                }

                // Save the workbook with the protected VBA project
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
