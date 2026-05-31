using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsMacroCopyExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template workbook that contains the macro which creates charts
            string templatePath = "TemplateWithMacro.xlsm";

            // Load the template workbook (macro-enabled)
            Workbook templateWorkbook = new Workbook(templatePath);

            // List of target workbook file names (output files)
            List<string> targetFiles = new List<string>
            {
                "Target1.xlsm",
                "Target2.xlsm",
                "Target3.xlsm"
                // Add more target file names as needed
            };

            // Loop through each target workbook, copy the VBA project (macro) from the template,
            // and save the result as a macro‑enabled workbook.
            foreach (string targetPath in targetFiles)
            {
                // Create an empty workbook
                Workbook targetWorkbook = new Workbook();

                // Enable macros for the target workbook (optional but clarifies intent)
                targetWorkbook.Settings.EnableMacros = true;

                // Copy the entire VBA project (including modules, references, etc.) from the template
                targetWorkbook.VbaProject.Copy(templateWorkbook.VbaProject);

                // Save the target workbook as a macro‑enabled file
                targetWorkbook.Save(targetPath, SaveFormat.Xlsm);
            }

            Console.WriteLine("Macro copied to all target workbooks successfully.");
        }
    }
}