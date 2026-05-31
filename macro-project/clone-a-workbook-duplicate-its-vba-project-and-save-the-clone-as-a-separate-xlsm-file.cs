using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace WorkbookCloneApp
{
    class Program
    {
        static void Main()
        {
            const string sourcePath = "source.xlsm";
            const string clonePath = "clone.xlsm";

            try
            {
                // Ensure the source file exists before loading
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook that contains VBA macros
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create an empty workbook that will hold the clone
                Workbook clonedWorkbook = new Workbook();

                // Configure copy options to retain macros/VBA project
                CopyOptions copyOptions = new CopyOptions
                {
                    KeepMacros = true
                };

                // Copy all contents from the source workbook to the clone
                sourceWorkbook.Copy(clonedWorkbook, copyOptions);

                // Explicitly duplicate the VBA project (optional, ensures full copy)
                clonedWorkbook.VbaProject.Copy(sourceWorkbook.VbaProject);

                // Save the cloned workbook as a macro‑enabled file
                clonedWorkbook.Save(clonePath, SaveFormat.Xlsm);

                Console.WriteLine($"Workbook cloned successfully to {clonePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}