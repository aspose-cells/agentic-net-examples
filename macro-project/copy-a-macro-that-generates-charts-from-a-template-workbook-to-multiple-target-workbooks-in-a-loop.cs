using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains the macro (must be a macro‑enabled file)
            string templatePath = "TemplateWithMacro.xlsm";
            Workbook templateWorkbook = new Workbook(templatePath);

            // List of target workbook file names (can be any supported format)
            string[] targetFiles = { "Target1.xlsx", "Target2.xlsx", "Target3.xlsx" };

            foreach (string targetFile in targetFiles)
            {
                // Create an empty workbook for the target
                Workbook targetWorkbook = new Workbook();

                // Enable macros in the target workbook settings (optional but safe)
                targetWorkbook.Settings.EnableMacros = true;

                // Copy the VBA project (all modules, references, etc.) from the template to the target
                targetWorkbook.VbaProject.Copy(templateWorkbook.VbaProject);

                // Save the target workbook as a macro‑enabled file
                string outputPath = Path.ChangeExtension(targetFile, ".xlsm");
                targetWorkbook.Save(outputPath, SaveFormat.Xlsm);

                Console.WriteLine($"Macro copied to: {outputPath}");
            }
        }
    }
}