// Title: Export VBA modules from an Excel workbook to individual .bas files using Aspose.Cells for .NET
// Common Searches: C# code to extract VBA modules from an .xlsx file with Aspose.Cells | how to save each VBA module as a separate .bas file in .NET | Aspose.Cells export VBA project modules for version control | write VBA module code to files using Aspose.Cells VbaModule class | automate VBA module extraction from Excel workbook using C#
// Tags: Aspose.Cells VBA module extraction | C# write VBA module to .bas | export .bas files from Excel workbook | version control VBA code with Aspose | VbaModule Codes property usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba; // Required for VbaModule

namespace VbaExportExample
{
    // The example loads an Excel workbook with Aspose.Cells, checks for a VBA project, iterates through each VbaModule, and writes the module's source code to uniquely named .bas files in a specified folder, creating the folder if needed and handling errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook that contains VBA modules
            string workbookPath = @"C:\Path\To\YourWorkbook.xlsx";

            // Folder where the exported .bas files will be saved
            string exportFolder = @"C:\Path\To\ExportedVbaModules";

            try
            {
                // Verify workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook not found at '{workbookPath}'.");
                    return;
                }

                // Ensure the export folder exists
                if (!Directory.Exists(exportFolder))
                {
                    Directory.CreateDirectory(exportFolder);
                }

                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(workbookPath);

                // Check if the workbook contains a VBA project
                if (workbook.VbaProject == null || workbook.VbaProject.Modules.Count == 0)
                {
                    Console.WriteLine("No VBA modules found in the workbook.");
                    return;
                }

                // Iterate through each VBA module and export its code to a .bas file
                foreach (VbaModule module in workbook.VbaProject.Modules)
                {
                    // Retrieve the VBA source code from the module
                    string vbaCode = module.Codes;

                    // Construct a safe file name for the .bas file
                    string safeModuleName = string.Join("_", module.Name.Split(Path.GetInvalidFileNameChars()));
                    string basFilePath = Path.Combine(exportFolder, $"{safeModuleName}.bas");

                    // Write the VBA code to the .bas file
                    File.WriteAllText(basFilePath, vbaCode);

                    Console.WriteLine($"Exported module '{module.Name}' to '{basFilePath}'.");
                }

                Console.WriteLine("All VBA modules have been exported.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
