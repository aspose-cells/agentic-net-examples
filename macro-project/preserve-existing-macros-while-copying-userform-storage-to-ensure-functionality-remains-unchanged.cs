using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroCopy
{
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
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string sourcePath = "SourceWithMacros.xlsm";
            const string destPath = "DestinationWithMacrosAndUserForms.xlsm";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

            // Load the source workbook that contains macros and userforms
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create an empty destination workbook
            Workbook destWorkbook = new Workbook();

            // Configure copy options to keep macros during the copy operation
            CopyOptions copyOptions = new CopyOptions
            {
                KeepMacros = true
            };

            // Copy the entire workbook (worksheets, data, macros, etc.) to the destination
            destWorkbook.Copy(sourceWorkbook, copyOptions);

            // Copy VBA references (libraries) from source to destination
            destWorkbook.VbaProject.References.Copy(sourceWorkbook.VbaProject.References);

            // Iterate through all VBA modules in the source workbook
            foreach (VbaModule srcModule in sourceWorkbook.VbaProject.Modules)
            {
                // Retrieve the designer storage for the current module (null for non‑UserForm modules)
                byte[] designerStorage = sourceWorkbook.VbaProject.Modules.GetDesignerStorage(srcModule.Name);

                // If designer storage exists, the module is a UserForm; copy it to the destination
                if (designerStorage != null && designerStorage.Length > 0)
                {
                    // Add the UserForm to the destination VBA project with its code and designer data
                    destWorkbook.VbaProject.Modules.AddUserForm(
                        srcModule.Name,   // UserForm name
                        srcModule.Codes,  // VBA code associated with the UserForm
                        designerStorage   // Binary designer storage (.frx data)
                    );
                }
            }

            // Save the destination workbook; macros and UserForms are preserved
            destWorkbook.Save(destPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{destPath}'.");
        }
    }
}