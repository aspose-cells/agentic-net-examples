using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaExport
{
    public class VbaExporter
    {
        // Exports VBA modules from all macro-enabled workbooks in a folder to .bas files.
        public static void Run(string inputFolder, string outputFolder)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Get all Excel macro-enabled files (*.xlsm) in the input folder.
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsm", SearchOption.TopDirectoryOnly);

            foreach (string workbookPath in workbookFiles)
            {
                // Verify the workbook file still exists.
                if (!File.Exists(workbookPath))
                {
                    continue;
                }

                try
                {
                    // Load the workbook (lifecycle rule: load).
                    Workbook workbook = new Workbook(workbookPath);

                    // Access the VBA project; it may be null if the workbook has no macros.
                    VbaProject vbaProject = workbook.VbaProject;
                    if (vbaProject == null)
                    {
                        // No VBA project present; skip this workbook.
                        continue;
                    }

                    // Iterate through all VBA modules in the project.
                    for (int i = 0; i < vbaProject.Modules.Count; i++)
                    {
                        VbaModule module = vbaProject.Modules[i];

                        // Retrieve the VBA source code from the module.
                        string code = module.Codes ?? string.Empty;

                        // Construct a file name: <WorkbookName>_<ModuleName>.bas
                        string workbookName = Path.GetFileNameWithoutExtension(workbookPath);
                        string safeModuleName = MakeFileSystemSafe(module.Name);
                        string basFileName = $"{workbookName}_{safeModuleName}.bas";

                        // Full path for the .bas file.
                        string basFilePath = Path.Combine(outputFolder, basFileName);

                        // Write the code to the .bas file (lifecycle rule: save).
                        File.WriteAllText(basFilePath, code);
                    }
                }
                catch (Exception ex)
                {
                    // Log and continue with next workbook.
                    Console.Error.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
                }
            }
        }

        // Helper to replace invalid filename characters.
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }

    // Entry point for the console application.
    public static class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: AsposeCellsVbaExport <inputFolder> <outputFolder>");
                    return;
                }

                string inputFolder = args[0];
                string outputFolder = args[1];

                if (!Directory.Exists(inputFolder))
                {
                    Console.Error.WriteLine($"Input folder does not exist: {inputFolder}");
                    return;
                }

                VbaExporter.Run(inputFolder, outputFolder);
                Console.WriteLine("VBA export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}