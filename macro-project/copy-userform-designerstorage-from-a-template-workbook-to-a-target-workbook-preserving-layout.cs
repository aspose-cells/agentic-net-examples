using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CopyUserFormDesignerStorage
{
    static void Main()
    {
        try
        {
            // Paths for the template, target, and output workbooks
            string templatePath = "TemplateWithUserForm.xlsm";
            string targetPath = "TargetWorkbook.xlsm";
            string outputPath = "TargetWorkbook_WithUserForm.xlsm";

            // Verify that the template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook that contains the UserForm
            Workbook templateWb = new Workbook(templatePath);

            // Load the target workbook if it exists; otherwise create a new workbook
            Workbook targetWb = File.Exists(targetPath) ? new Workbook(targetPath) : new Workbook();

            // Name of the UserForm to copy (adjust if different)
            string userFormName = "UserForm1";

            // ----- Retrieve the designer storage (binary .frx data) from the template -----
            byte[] designerStorage = templateWb.VbaProject.Modules.GetDesignerStorage(userFormName);

            // ----- Locate the module that holds the UserForm's VBA code -----
            int sourceModuleIndex = -1;
            for (int i = 0; i < templateWb.VbaProject.Modules.Count; i++)
            {
                if (templateWb.VbaProject.Modules[i].Name.Equals(userFormName, StringComparison.OrdinalIgnoreCase))
                {
                    sourceModuleIndex = i;
                    break;
                }
            }

            if (sourceModuleIndex == -1)
            {
                Console.WriteLine($"UserForm '{userFormName}' not found in the template workbook.");
                return;
            }

            // Extract the VBA code (the .bas part) from the source module
            string userFormCode = templateWb.VbaProject.Modules[sourceModuleIndex].Codes;

            // ----- Add the UserForm to the target workbook's VBA project -----
            int addedIndex = targetWb.VbaProject.Modules.AddUserForm(userFormName, userFormCode, designerStorage);
            Console.WriteLine($"UserForm added to target workbook at module index: {addedIndex}");

            // Save the target workbook (preserving macros and the newly added UserForm)
            targetWb.Save(outputPath);
            Console.WriteLine($"Target workbook saved successfully as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}