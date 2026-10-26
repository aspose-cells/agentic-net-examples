// Title: Copy VBA UserForm code from a template Excel workbook to another workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to copy VBA UserForm code between Excel files with Aspose.Cells in C# | Aspose.Cells example for transferring VBA modules from a template workbook to a target workbook | C# copy VBA project from one workbook to another while preserving UserForm layout using Aspose.Cells | move VBA UserForm source code to another Excel workbook with Aspose.Cells .NET API | Aspose.Cells VBA migration tutorial for Excel workbooks in C#
// Tags: Aspose.Cells copy VBA modules C# | VbaModule source code transfer Aspose.Cells | Excel workbook VBA project migration .NET | UserForm code copy limitation Aspose.Cells | Aspose.Cells VBA handling in C#

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The sample loads a template and a target Excel workbook, checks that both contain VBA projects, then iterates through each VbaModule in the template. For each module it either finds a matching module in the target or adds a new one, and copies the VBA source code. Designer storage for UserForms is not exposed via the public API, so only the code is transferred. Finally, the updated target workbook is saved.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string targetPath = "Target.xlsx";
            const string outputPath = "TargetWithUserForm.xlsx";

            // Verify that the input files exist to avoid FileNotFoundException
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");
            if (!File.Exists(targetPath))
                throw new FileNotFoundException($"Target file not found: {targetPath}");

            // Load the workbooks
            Workbook templateWb = new Workbook(templatePath);
            Workbook targetWb = new Workbook(targetPath);

            // Ensure the template workbook contains a VBA project
            if (templateWb.VbaProject == null)
                throw new InvalidOperationException("Template workbook does not contain a VBA project.");

            // Ensure the target workbook contains a VBA project
            if (targetWb.VbaProject == null)
                throw new InvalidOperationException("Target workbook does not contain a VBA project.");

            // Copy each VBA module from the template to the target workbook
            foreach (VbaModule tmplModule in templateWb.VbaProject.Modules)
            {
                // Find a module with the same name in the target workbook
                VbaModule targetModule = targetWb.VbaProject.Modules
                    .FirstOrDefault(m => m.Name == tmplModule.Name);

                // If not present, add a new module with the same type
                if (targetModule == null)
                {
                    // Modules.Add returns the index of the newly added module
                    int newIndex = targetWb.VbaProject.Modules.Add(tmplModule.Type, tmplModule.Name);
                    targetModule = targetWb.VbaProject.Modules[newIndex];
                }

                // Copy the VBA source code
                targetModule.Codes = tmplModule.Codes;
                // Note: Designer data for UserForms is not exposed via the public API,
                // so only the source code is copied.
            }

            // Save the modified target workbook
            targetWb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
