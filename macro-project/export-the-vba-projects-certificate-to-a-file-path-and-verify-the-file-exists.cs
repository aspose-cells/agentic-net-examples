// Title: Export a VBA project's certificate from an .xlsm workbook to a .vba file and confirm the file exists using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a macro‑enabled workbook with Aspose.Cells, exports its VBA project (including the digital certificate) to a given .vba file, and creates the destination folder if it does not exist. | After calling ExportVbaProject, add logic that uses System.IO.File.Exists to verify the .vba file was created and output a success message. | Include try‑catch handling to report any errors that occur during workbook loading, VBA export, or file‑system operations.
// Common Searches: how to export VBA project certificate from .xlsm using Aspose.Cells C# | C# Aspose.Cells ExportVbaProject example with file existence verification | save VBA project to .vba file and check export result in .NET
// Tags: Aspose.Cells ExportVbaProject to .vba | C# check file existence after VBA export | validate VbaProject presence in workbook | ensure export directory exists .NET | export VBA project certificate using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba; // Required for VBA related extensions

namespace ExportVbaCertificate
{
    // The program loads a macro‑enabled .xlsm workbook, confirms it contains a VBA project, creates the target directory if needed, exports the VBA project (including its digital certificate) to a .vba file via Aspose.Cells, and then verifies that the exported file exists.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file that contains a VBA project
            string workbookPath = @"C:\Path\To\YourWorkbook.xlsm";

            // Path where the exported VBA project (including its certificate) will be saved
            string vbaProjectExportPath = @"C:\Path\To\ExportedVbaProject.vba";

            try
            {
                // Ensure the source workbook exists before loading
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Verify that the workbook contains a VBA project
                if (workbook.VbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Ensure the export directory exists
                string exportDir = Path.GetDirectoryName(vbaProjectExportPath);
                if (!string.IsNullOrEmpty(exportDir) && !Directory.Exists(exportDir))
                {
                    Directory.CreateDirectory(exportDir);
                }

                // Export the VBA project (including its digital certificate, if present)
                // Use dynamic to avoid compile‑time binding issues with different library versions
                dynamic vbaProject = workbook.VbaProject;
                vbaProject.ExportVbaProject(vbaProjectExportPath);

                // Confirm that the export file now exists
                bool fileExists = File.Exists(vbaProjectExportPath);
                Console.WriteLine($"VBA project exported: {fileExists}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
