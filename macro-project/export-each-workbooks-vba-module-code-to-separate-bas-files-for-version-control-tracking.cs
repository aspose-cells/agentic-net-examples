using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace VbaExportExample
{
    class Program
    {
        // Replace characters that are invalid in file names
        private static string MakeValidFileName(string name)
        {
            string invalidChars = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            string escaped = Regex.Escape(invalidChars);
            string pattern = $"[{escaped}]";
            return Regex.Replace(name, pattern, "_");
        }

        static void Main(string[] args)
        {
            // Folder containing the workbooks to process
            string inputFolder = @"C:\Workbooks";
            // Folder where the .bas files will be written
            string outputFolder = @"C:\VbaExports";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all macro-enabled Excel files in the input folder
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsm");

            foreach (string workbookPath in workbookFiles)
            {
                // Load the workbook (lifecycle: load)
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project; it may be null if the workbook has no macros
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null || vbaProject.Modules.Count == 0)
                {
                    continue; // No VBA modules to export
                }

                // Iterate through each VBA module in the project
                foreach (VbaModule module in vbaProject.Modules)
                {
                    // Retrieve the VBA source code
                    string code = module.Codes;
                    if (string.IsNullOrEmpty(code))
                    {
                        continue; // Skip empty modules
                    }

                    // Build a safe file name: WorkbookName_ModuleName.bas
                    string workbookBaseName = Path.GetFileNameWithoutExtension(workbookPath);
                    string safeModuleName = MakeValidFileName(module.Name);
                    string basFileName = $"{workbookBaseName}_{safeModuleName}.bas";

                    // Full path for the .bas file
                    string basFilePath = Path.Combine(outputFolder, basFileName);

                    // Write the code to the .bas file (lifecycle: save)
                    File.WriteAllText(basFilePath, code);
                }
            }

            Console.WriteLine("VBA module export completed.");
        }
    }
}