using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CopyMacrosDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                const string sourcePath = "source_with_macros.xlsm";
                const string destPath = "destination_with_macros.xlsm";

                // Verify source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook that contains macros (must be a macro-enabled file)
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create an empty destination workbook
                Workbook destWorkbook = new Workbook();

                // Preserve the macro security setting from the source workbook
                destWorkbook.Settings.EnableMacros = sourceWorkbook.Settings.EnableMacros;

                // Set copy options to keep macros during the copy operation
                CopyOptions options = new CopyOptions
                {
                    KeepMacros = true
                };

                // Copy the entire source workbook (including macros) into the destination workbook
                sourceWorkbook.Copy(destWorkbook, options);

                // Save the destination workbook as a macro-enabled file
                destWorkbook.Save(destPath);
                Console.WriteLine($"Destination workbook saved successfully: {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}