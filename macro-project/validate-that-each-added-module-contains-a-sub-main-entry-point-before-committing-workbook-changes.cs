using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaValidation
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (macro‑enabled format will be used on save)
                Workbook workbook = new Workbook();

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Add two class modules
                int idx1 = vbaProject.Modules.Add(VbaModuleType.Class, "ModuleOne");
                VbaModule moduleOne = vbaProject.Modules[idx1];
                moduleOne.Codes = @"Sub Main()
    MsgBox ""Hello from ModuleOne!""
End Sub";

                int idx2 = vbaProject.Modules.Add(VbaModuleType.Class, "ModuleTwo");
                VbaModule moduleTwo = vbaProject.Modules[idx2];
                // This module intentionally lacks Sub Main to demonstrate validation
                moduleTwo.Codes = @"Sub Helper()
    MsgBox ""Helper routine""
End Sub";

                // Validate that each *class* module contains a Sub Main entry point
                foreach (VbaModule module in vbaProject.Modules)
                {
                    // Skip non‑class modules such as ThisWorkbook or Sheet modules
                    if (module.Type != VbaModuleType.Class)
                        continue;

                    bool hasMain = Regex.IsMatch(
                        module.Codes ?? string.Empty,
                        @"^\s*Sub\s+Main\s*\(\s*\)",
                        RegexOptions.IgnoreCase | RegexOptions.Multiline);

                    if (!hasMain)
                    {
                        throw new InvalidOperationException(
                            $"VBA module \"{module.Name}\" does not contain a Sub Main entry point.");
                    }
                }

                // All modules passed validation – save the workbook as a macro‑enabled file
                string outputPath = "ValidatedWorkbook.xlsm";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log any errors without crashing the application
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}