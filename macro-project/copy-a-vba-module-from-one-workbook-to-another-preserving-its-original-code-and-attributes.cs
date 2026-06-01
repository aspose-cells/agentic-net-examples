using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsDemo
{
    public class CopyVbaModuleDemo
    {
        public static void Run()
        {
            try
            {
                const string sourcePath = "source.xlsm";
                const string destPath = "dest.xlsm";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook that already contains the VBA module
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Ensure the source workbook has a VBA project with at least one module
                if (sourceWorkbook.VbaProject == null || sourceWorkbook.VbaProject.Modules.Count == 0)
                {
                    Console.WriteLine("Source workbook does not contain any VBA modules.");
                    return;
                }

                // Access the first VBA module in the source workbook
                VbaModule sourceModule = sourceWorkbook.VbaProject.Modules[0];

                // Create a new (empty) destination workbook
                Workbook destWorkbook = new Workbook();

                // Add a new module to the destination workbook with the same type and name as the source module
                int destModuleIndex = destWorkbook.VbaProject.Modules.Add(sourceModule.Type, sourceModule.Name);
                VbaModule destModule = destWorkbook.VbaProject.Modules[destModuleIndex];

                // Copy the VBA code from the source module to the destination module
                destModule.Codes = sourceModule.Codes;

                // Save the destination workbook as a macro‑enabled file to preserve the VBA project
                destWorkbook.Save(destPath, SaveFormat.Xlsm);
                Console.WriteLine($"Destination workbook saved successfully: {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CopyVbaModuleDemo.Run();
        }
    }
}