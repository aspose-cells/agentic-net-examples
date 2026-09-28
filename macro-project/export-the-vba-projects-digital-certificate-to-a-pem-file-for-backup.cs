// Title: Export a VBA project's digital certificate from an Excel workbook to a .vba file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, verifies the presence of a VbaProject, and uses reflection to invoke ExportVbaProject to save the project (including its embedded certificate) to a .vba file, with proper handling for missing methods. | Generate a C# snippet that checks the input file path, extracts the VBA project from a workbook, and writes the project and its digital signature to a specified .vba backup file, incorporating comprehensive exception handling.
// Common Searches: how to save VBA project with digital certificate from Excel using Aspose.Cells C# | Aspose.Cells export VBA project to .vba file via reflection when ExportVbaProject is not directly accessible | C# backup VBA project from .xlsx workbook including its certificate | retrieve embedded VBA digital signature from workbook using Aspose.Cells .NET | export VBA project from workbook to file with error handling Aspose.Cells
// Tags: Aspose.Cells export VBA project to .vba | C# extract VBA digital certificate from workbook | reflection invoke ExportVbaProject method | backup VBA project Aspose.Cells .NET | check VBA project existence workbook Aspose

using Aspose.Cells;
using Aspose.Cells.Vba;
using System;
using System.IO;
using System.Reflection;

// The example loads an Excel workbook, confirms a VBA project is present, and uses reflection to call ExportVbaProject (if available) to export the VBA project—including its embedded digital certificate—to a .vba file, while handling missing files and method availability errors.
class ExportVbaCertificate
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx"; // TODO: replace with your workbook path

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that may contain a VBA project
            Workbook workbook = new Workbook(inputPath);

            // Check if the workbook contains a VBA project
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("No VBA project found in the workbook.");
                return;
            }

            // Export the VBA project (including its digital certificate) to a file
            string vbaExportPath = "vba_project.vba"; // TODO: replace with desired output path

            try
            {
                // Use reflection to call ExportVbaProject if it exists in the current Aspose.Cells version
                MethodInfo exportMethod = vbaProject.GetType().GetMethod(
                    "ExportVbaProject",
                    new[] { typeof(string) });

                if (exportMethod != null)
                {
                    exportMethod.Invoke(vbaProject, new object[] { vbaExportPath });
                    Console.WriteLine($"VBA project exported to: {vbaExportPath}");
                }
                else
                {
                    Console.WriteLine("ExportVbaProject method is not available in this Aspose.Cells version.");
                }
            }
            catch (TargetInvocationException tie)
            {
                Console.WriteLine($"Failed to export VBA project: {tie.InnerException?.Message ?? tie.Message}");
            }
            catch (Exception exportEx)
            {
                Console.WriteLine($"Failed to export VBA project: {exportEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
