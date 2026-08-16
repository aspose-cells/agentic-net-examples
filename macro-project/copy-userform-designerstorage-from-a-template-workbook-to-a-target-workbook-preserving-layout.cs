// Title: Copy UserForm Designer Storage Between Macro‑Enabled Workbooks with Aspose.Cells for .NET
// Description: Shows how to load a macro‑enabled template, copy its worksheets, theme and VBA macros, then transfer each UserForm's code and .frx designer storage to a new workbook using Aspose.Cells VbaProject methods, and finally save the result as an .xlsm file.
// Keywords: Aspose.Cells | C# | copy UserForm | VBA designer storage | .frx | macro‑enabled workbook | VbaProject | AddUserForm | GetDesignerStorage | Excel automation | workbook cloning
// Common Searches: Aspose.Cells copy UserForm between workbooks | transfer VBA designer storage .frx with C# | preserve macros when cloning Excel file Aspose.Cells | how to add UserForm to a new workbook programmatically | copy macro‑enabled template to new workbook using Aspose
// Developer Intent: Programmatically duplicate a UserForm—including its VBA code and .frx designer storage—from a template workbook to another workbook while retaining all macros and layout.
// Use Cases: Create client‑specific macro workbooks from a standard template that contains pre‑designed UserForms. | Migrate legacy Excel files with embedded UserForms into a new project structure without losing form design. | Automate generation of reporting workbooks that need identical UserForm interfaces across multiple files.
// AI Prompts: Write C# code using Aspose.Cells to copy all UserForm designer storage from a source .xlsm to a destination .xlsm, preserving VBA code and macros. | Explain the role of VbaProject.Modules.GetDesignerStorage and VbaProject.Modules.AddUserForm in transferring UserForms between workbooks. | Suggest robust error‑handling patterns for copying VBA projects and UserForms with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Shows how to load a macro‑enabled template, copy its worksheets, theme and VBA macros, then transfer each UserForm's code and .frx designer storage to a new workbook using Aspose.Cells VbaProject methods, and finally save the result as an .xlsm file.
    class CopyUserFormDesignerStorage
    {
        static void Main()
        {
            try
            {
                const string templatePath = "TemplateWithUserForm.xlsm";
                const string targetPath = "TargetWithCopiedUserForm.xlsm";

                // Verify that the template file exists to avoid FileNotFoundException
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {Path.GetFullPath(templatePath)}");
                    return;
                }

                // Load the template workbook (contains the UserForm and its designer storage)
                Workbook templateWb = new Workbook(templatePath);

                // Create an empty workbook that will receive the copied content
                Workbook targetWb = new Workbook();

                // -----------------------------------------------------------------
                // 1. Copy worksheets, theme and macros from the template to the target
                // -----------------------------------------------------------------
                CopyOptions copyOpts = new CopyOptions
                {
                    KeepMacros = true // preserve VBA macros
                };
                targetWb.Copy(templateWb, copyOpts); // copies worksheets, theme, etc.

                // -----------------------------------------------------------------
                // 2. Copy the UserForm designer storage (and its code) manually
                // -----------------------------------------------------------------
                VbaProject templateVba = templateWb.VbaProject;
                VbaProject targetVba = targetWb.VbaProject;

                // Ensure both VBA projects are available
                if (templateVba == null || targetVba == null)
                {
                    Console.WriteLine("One of the workbooks does not contain a VBA project.");
                    return;
                }

                // Iterate through all modules in the template VBA project
                foreach (VbaModule tmplModule in templateVba.Modules)
                {
                    // We are interested only in Designer modules (UserForms)
                    if (tmplModule.Type == VbaModuleType.Designer)
                    {
                        string formName = tmplModule.Name;          // e.g., "UserForm1"
                        string formCode = tmplModule.Codes;         // VBA code behind the form

                        // Retrieve the binary designer storage (.frx) for this form
                        byte[] designerStorage = templateVba.Modules.GetDesignerStorage(formName);

                        // Add the UserForm to the target workbook's VBA project
                        targetVba.Modules.AddUserForm(formName, formCode, designerStorage);
                    }
                }

                // -----------------------------------------------------------------
                // 3. Save the resulting workbook (still macro‑enabled)
                // -----------------------------------------------------------------
                targetWb.Save(targetPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(targetPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
