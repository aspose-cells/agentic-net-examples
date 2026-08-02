// Title: Copy an XLSM workbook with macros and UserForm storage using Aspose.Cells for .NET
// Description: Demonstrates how to load a macro‑enabled workbook, copy its worksheets, VBA modules, references, and UserForm designer storage, and save the result as an Xlsm file. The example uses CopyOptions.KeepMacros and VbaProject.Modules.AddUserForm to ensure all macros and UserForms remain functional after the copy.
// Keywords: Aspose.Cells copy macros | preserve UserForm storage | CopyOptions KeepMacros .NET | VbaProject AddUserForm | duplicate XLSM workbook | C# Aspose.Cells VBA modules | macro‑enabled workbook copy
// Common Searches: copy xlsm file with macros Aspose.Cells | preserve VBA UserForm when duplicating workbook C# | Aspose.Cells KeepMacros example | how to copy VBA references with Aspose.Cells | add UserForm designer storage Aspose.Cells .NET
// Developer Intent: Copy a macro‑enabled Excel workbook while keeping all VBA code, references, and UserForm designer storage intact.
// Use Cases: Create personalized copies of a template workbook that retain existing macros and UserForms. | Migrate legacy Excel macros to a new file without losing UserForm layouts or external VBA references. | Automate batch generation of macro‑enabled reports that require identical VBA functionality.
// AI Prompts: Write C# code with Aspose.Cells to duplicate an XLSM workbook and preserve VBA modules, references, and UserForm designer storage. | Explain the role of VbaProject.Modules.AddUserForm and how to retrieve designer storage for a UserForm in Aspose.Cells. | Suggest robust error‑handling patterns for copying macro‑enabled workbooks using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroCopyDemo
{
    // Demonstrates how to load a macro‑enabled workbook, copy its worksheets, VBA modules, references, and UserForm designer storage, and save the result as an Xlsm file. The example uses CopyOptions.KeepMacros and VbaProject.Modules.AddUserForm to ensure all macros and UserForms remain functional after the copy.
    public class PreserveMacrosAndUserForms
    {
        // Entry point required for console application
        public static void Main(string[] args)
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
            const string sourcePath = "source_with_userforms.xlsm";
            const string destPath = "destination_preserved.xlsm";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the source workbook that contains macros and UserForms
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create an empty destination workbook
                Workbook destinationWorkbook = new Workbook();

                // Configure copy options to keep macros during the workbook copy
                CopyOptions copyOptions = new CopyOptions
                {
                    KeepMacros = true
                };

                // Copy the entire workbook (worksheets, data, macros) while preserving macros
                destinationWorkbook.Copy(sourceWorkbook, copyOptions);

                // Copy VBA references (e.g., external libraries) from source to destination
                if (sourceWorkbook.VbaProject != null && destinationWorkbook.VbaProject != null)
                {
                    destinationWorkbook.VbaProject.References.Copy(sourceWorkbook.VbaProject.References);
                }

                // Iterate through each VBA module in the source workbook
                for (int i = 0; i < sourceWorkbook.VbaProject.Modules.Count; i++)
                {
                    VbaModule srcModule = sourceWorkbook.VbaProject.Modules[i];

                    // Obtain designer storage for the module (non‑null for UserForms)
                    byte[] designerStorage = sourceWorkbook.VbaProject.Modules.GetDesignerStorage(srcModule.Name);

                    if (designerStorage != null && designerStorage.Length > 0)
                    {
                        // The module is a UserForm – add it to the destination VBA project
                        destinationWorkbook.VbaProject.Modules.AddUserForm(srcModule.Name, srcModule.Codes, designerStorage);
                    }
                    else
                    {
                        // Regular VBA module – add it and copy its code
                        int newIndex = destinationWorkbook.VbaProject.Modules.Add(srcModule.Type, srcModule.Name);
                        destinationWorkbook.VbaProject.Modules[newIndex].Codes = srcModule.Codes;
                    }
                }

                // Save the destination workbook as a macro‑enabled file
                destinationWorkbook.Save(destPath, SaveFormat.Xlsm);

                Console.WriteLine("Workbook copied with macros and UserForms preserved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An exception occurred during processing: {ex.Message}");
            }
        }
    }
}
