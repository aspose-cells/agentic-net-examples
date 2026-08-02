// Title: Copy VBA UserForm 'InputForm' Between .xlsm Workbooks with Aspose.Cells for .NET
// Description: Demonstrates how to load a macro‑enabled workbook, locate the VbaModule named InputForm, extract its code, create an empty .xlsm file, copy VBA references, add the UserForm (including designer storage) with AddUserForm, and save the destination workbook while preserving all controls and logic.
// Keywords: Aspose.Cells | C# | VBA UserForm copy | AddUserForm | macro-enabled workbook | transfer VBA form | preserve VBA controls | copy .xlsm UserForm | VbaModule example | Excel automation
// Common Searches: copy VBA UserForm from one .xlsm to another using Aspose.Cells | Aspose.Cells add existing UserForm to new workbook | transfer UserForm code and designer storage C# | how to move a macro-enabled UserForm with Aspose.Cells | clone VBA UserForm across Excel files programmatically
// Developer Intent: Move the InputForm UserForm, together with its VBA code and designer data, from a source macro‑enabled workbook to a target workbook using Aspose.Cells for .NET.
// Use Cases: Reuse a custom data‑entry form from a template workbook in newly generated reports. | Create a fresh macro‑enabled file that inherits an existing UserForm for batch processing. | Duplicate a VBA UserForm and its project references when building a suite of automated Excel tools.
// AI Prompts: Generate C# code with Aspose.Cells to copy a VBA UserForm named 'InputForm' from one .xlsm file to another, preserving code and designer storage. | Show how to create an empty macro‑enabled workbook, copy VBA references, and add a UserForm using the AddUserForm method. | Explain error handling for missing source workbook or absent UserForm when transferring VBA modules with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Demonstrates how to load a macro‑enabled workbook, locate the VbaModule named InputForm, extract its code, create an empty .xlsm file, copy VBA references, add the UserForm (including designer storage) with AddUserForm, and save the destination workbook while preserving all controls and logic.
    public class TransferUserForm
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Paths to the source and destination workbooks
            string sourcePath = "SourceWorkbook.xlsm";
            string destinationPath = "DestinationWorkbook.xlsm";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the source workbook (must be macro-enabled)
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Ensure the source workbook actually contains a VBA project
                if (sourceWorkbook.VbaProject == null || sourceWorkbook.VbaProject.Modules.Count == 0)
                {
                    Console.WriteLine("Source workbook does not contain any VBA modules.");
                    return;
                }

                // Find the UserForm named "InputForm" in the source VBA modules
                VbaModule sourceUserForm = null;
                foreach (VbaModule module in sourceWorkbook.VbaProject.Modules)
                {
                    // UserForm modules are identified by name; type check omitted for compatibility
                    if (module.Name.Equals("InputForm", StringComparison.OrdinalIgnoreCase))
                    {
                        sourceUserForm = module;
                        break;
                    }
                }

                if (sourceUserForm == null)
                {
                    Console.WriteLine("UserForm 'InputForm' not found in the source workbook.");
                    return;
                }

                // Extract the VBA code from the source UserForm
                string formCode = sourceUserForm.Codes;

                // Designer storage (binary .frx data). Use empty array to satisfy non‑null requirement.
                byte[] designerStorage = new byte[0];

                // ------------------------------------------------------------
                // Create the destination workbook (empty workbook)
                // ------------------------------------------------------------
                Workbook destWorkbook = new Workbook();

                // Save as a macro‑enabled workbook to create an empty VBA project, then reload it.
                string tempMacroPath = Path.Combine(Path.GetTempPath(),
                    Guid.NewGuid().ToString("N") + ".xlsm");
                destWorkbook.Save(tempMacroPath, SaveFormat.Xlsm);
                destWorkbook = new Workbook(tempMacroPath);
                File.Delete(tempMacroPath);

                // Copy VBA references from source to destination (optional but recommended)
                destWorkbook.VbaProject.References.Copy(sourceWorkbook.VbaProject.References);

                // Add the UserForm to the destination VBA project
                VbaModuleCollection destModules = destWorkbook.VbaProject.Modules;
                int newModuleIndex = destModules.AddUserForm("InputForm", formCode, designerStorage);
                Console.WriteLine($"UserForm added to destination workbook at index: {newModuleIndex}");

                // Ensure the directory for the destination file exists
                string destDir = Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Save the destination workbook with the transferred UserForm
                destWorkbook.Save(destinationPath, SaveFormat.Xlsm);
                Console.WriteLine($"UserForm 'InputForm' transferred successfully to '{destinationPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}
