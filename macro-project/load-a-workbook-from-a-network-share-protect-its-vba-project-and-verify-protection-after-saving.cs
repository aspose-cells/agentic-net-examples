// Title: Protect a VBA project in a macro-enabled workbook loaded from a UNC network share using Aspose.Cells for .NET and verify the protection after saving
// AI Prompts: Write C# code that opens an .xlsm file from a UNC path, uses Aspose.Cells to set a password on the workbook's VbaProject, saves the file, reloads it, and prints the VbaProject.IsProtected value. | Generate an Aspose.Cells example that demonstrates loading a macro-enabled workbook over a network share, applying VBA project protection with a password, persisting the changes, and confirming the protection status programmatically.
// Common Searches: aspnet load xlsm from network share and protect vba project with password | c# Aspose.Cells protect VBA project in macro-enabled workbook saved on UNC path | how to check if VBA project is protected after saving with Aspose.Cells | save protected macro workbook to network location using Aspose.Cells .NET
// Tags: apply password to VBA project Aspose.Cells | load XLSM from UNC path C# | save protected macro workbook to network share | check VbaProject.IsProtected status | Aspose.Cells VBA protection example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionDemo
{
    // // Loads an .xlsm workbook from a UNC network share, applies password protection to its VBA project via Aspose.Cells, saves the workbook back to the share, reloads it, and outputs the VBA project's IsProtected flag.
    class Program
    {
        static void Main()
        {
            // Path to the workbook on a network share (replace with a valid path for testing)
            string inputPath = @"\\Server\Share\input.xlsm";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook from the network location
                Workbook workbook = new Workbook(inputPath);

                // Protect the VBA project (lock for viewing = false, set a password)
                if (workbook.VbaProject != null)
                {
                    workbook.VbaProject.Protect(false, "VbaPassword123");
                }

                // Define the output path and ensure its directory exists
                string outputPath = @"\\Server\Share\output_protected.xlsm";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook back to the network share (macro-enabled format)
                workbook.Save(outputPath, SaveFormat.Xlsm);

                // Reload the saved workbook to verify protection status
                Workbook reloadedWorkbook = new Workbook(outputPath);
                VbaProject vbaProject = reloadedWorkbook.VbaProject;

                // Output verification results
                Console.WriteLine("VBA Project IsProtected: " + (vbaProject?.IsProtected.ToString() ?? "null"));
                // The IsLockedForViewing property is not available in all versions; omitted for compatibility.
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
