// Title: Copy a macro‑enabled workbook and retain a VBA UserForm (code + .frx) using Aspose.Cells for .NET
// Description: Loads a source .xlsm, copies it with CopyOptions.KeepMacros, extracts the specified UserForm's VBA code and designer storage, inserts the form into a new workbook, and saves the result as a macro‑enabled file.
// Keywords: Aspose.Cells copy workbook macros | preserve VBA UserForm .frx | AddUserForm C# | CopyOptions KeepMacros example | macro enabled Excel duplication .NET | VBA module transfer Aspose | Excel UserForm cloning C#
// Common Searches: Aspose.Cells copy .xlsm keep macros | how to copy VBA UserForm with Aspose.Cells | preserve .frx designer storage when copying workbook | C# copy macro enabled Excel file and retain UserForm | Aspose.Cells AddUserForm example
// Developer Intent: Duplicate a macro‑enabled Excel workbook while keeping a chosen VBA UserForm’s code and binary designer data intact.
// Use Cases: Create client‑specific reports from a template, preserving the UI form used for data entry. | Migrate legacy .xlsm files to a new project without losing custom UserForm interfaces. | Automate generation of multiple workbooks that must share the same VBA UserForm across deployments.
// AI Prompts: Write C# code that copies an .xlsm workbook, keeps all macros, and adds a specific UserForm with its .frx storage using Aspose.Cells. | Explain how to detect a missing UserForm module during workbook copy and provide fallback handling in Aspose.Cells. | Show how to copy only selected VBA modules, including UserForms, from one workbook to another with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace PreserveMacrosAndUserFormApp
{
    // Loads a source .xlsm, copies it with CopyOptions.KeepMacros, extracts the specified UserForm's VBA code and designer storage, inserts the form into a new workbook, and saves the result as a macro‑enabled file.
    class PreserveMacrosAndUserForm
    {
        static void Main()
        {
            try
            {
                string sourcePath = "source_with_userform.xlsm";
                string destPath = "destination_with_userform.xlsm";

                // Verify that the source file exists before loading
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook that contains macros and a UserForm
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create an empty destination workbook
                Workbook destinationWorkbook = new Workbook();

                // Configure copy options to keep macros during the copy operation
                CopyOptions copyOptions = new CopyOptions
                {
                    KeepMacros = true
                };

                // Copy the entire source workbook into the destination workbook,
                // preserving all macros (including standard modules)
                destinationWorkbook.Copy(sourceWorkbook, copyOptions);

                // Name of the UserForm to be copied
                string userFormName = "UserForm1";

                // Attempt to retrieve the VBA module for the specified UserForm
                VbaModule sourceModule = null;
                try
                {
                    sourceModule = sourceWorkbook.VbaProject.Modules[userFormName];
                }
                catch
                {
                    // Module not found; will be handled below
                }

                if (sourceModule != null && destinationWorkbook.VbaProject != null)
                {
                    // Retrieve the VBA code of the UserForm
                    string userFormCode = sourceModule.Codes;

                    // Retrieve the designer storage (binary .frx data) of the UserForm
                    byte[] designerStorage = sourceWorkbook.VbaProject.Modules.GetDesignerStorage(userFormName);

                    // Insert the UserForm into the destination workbook, preserving both code and designer storage
                    destinationWorkbook.VbaProject.Modules.AddUserForm(userFormName, userFormCode, designerStorage);
                }
                else
                {
                    Console.WriteLine($"UserForm '{userFormName}' not found in the source workbook.");
                }

                // Save the destination workbook as a macro‑enabled file
                destinationWorkbook.Save(destPath, SaveFormat.Xlsm);
                Console.WriteLine($"Destination workbook saved to: {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
