using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CopyUserFormDesignerStorage
{
    static void Main()
    {
        // Paths to the template (source) workbook and the target workbook.
        string templatePath = "Template.xlsm";
        string targetPath   = "Target.xlsx";
        string outputPath   = "Result.xlsm";

        // Load the source workbook that contains the UserForm.
        Workbook sourceWorkbook = new Workbook(templatePath);

        // Load (or create) the destination workbook.
        Workbook destWorkbook = new Workbook(targetPath);

        // -----------------------------------------------------------------
        // 1. Copy worksheets and keep macros (including VBA project structure).
        // -----------------------------------------------------------------
        CopyOptions copyOptions = new CopyOptions();
        copyOptions.KeepMacros = true;               // Preserve macros while copying.
        destWorkbook.Copy(sourceWorkbook, copyOptions);

        // -----------------------------------------------------------------
        // 2. Copy the DesignerStorage (FRX binary) of each UserForm.
        // -----------------------------------------------------------------
        VbaProject srcVba = sourceWorkbook.VbaProject;
        VbaProject dstVba = destWorkbook.VbaProject;

        // Iterate through all VBA modules in the source workbook.
        foreach (VbaModule srcModule in srcVba.Modules)
        {
            // Only process modules of type Designer (UserForms).
            if (srcModule.Type == VbaModuleType.Designer)
            {
                string formName = srcModule.Name;          // e.g., "UserForm1"
                string formCode = srcModule.Codes;         // VBA code behind the form.

                // Retrieve the binary designer storage (FRX data) for the form.
                byte[] designerStorage = srcVba.Modules.GetDesignerStorage(formName);

                // Add the UserForm to the destination workbook's VBA project.
                // The method returns the index of the newly added module (not used here).
                dstVba.Modules.AddUserForm(formName, formCode, designerStorage);
            }
        }

        // -----------------------------------------------------------------
        // 3. Save the resulting workbook (must be a macro‑enabled format).
        // -----------------------------------------------------------------
        destWorkbook.Save(outputPath, SaveFormat.Xlsm);

        Console.WriteLine("UserForm DesignerStorage copied successfully to: " + outputPath);
    }
}