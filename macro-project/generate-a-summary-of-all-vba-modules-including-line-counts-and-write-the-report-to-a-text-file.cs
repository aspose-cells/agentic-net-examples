using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class VbaModulesSummary
{
    static void Main()
    {
        // Load a macro-enabled workbook that contains VBA modules
        string workbookPath = "input.xlsm";
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;
        if (vbaProject == null)
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
            return;
        }

        // Get the collection of VBA modules
        VbaModuleCollection modules = vbaProject.Modules;

        // Write the summary to a text file
        using (StreamWriter writer = new StreamWriter("VbaModulesReport.txt"))
        {
            writer.WriteLine($"Total Modules: {modules.Count}");
            writer.WriteLine();

            for (int i = 0; i < modules.Count; i++)
            {
                VbaModule module = modules[i];
                string code = module.Codes ?? string.Empty;

                // Count lines in the module code
                int lineCount = 0;
                if (!string.IsNullOrEmpty(code))
                {
                    lineCount = code.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).Length;
                }

                writer.WriteLine($"Module {i + 1}:");
                writer.WriteLine($"Name: {module.Name}");
                writer.WriteLine($"Type: {module.Type}");
                writer.WriteLine($"Line Count: {lineCount}");
                writer.WriteLine();
            }
        }

        Console.WriteLine("VBA modules summary has been written to VbaModulesReport.txt");
    }
}