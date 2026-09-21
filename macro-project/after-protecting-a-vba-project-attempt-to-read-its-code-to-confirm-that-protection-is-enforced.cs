// Title: Protect an Excel VBA project with a password using Aspose.Cells for .NET and confirm protection by attempting to read the macro code
// AI Prompts: Open an existing .xlsm file with Aspose.Cells, call Workbook.VbaProject.Protect(true, "myPassword"), and save the workbook as a new protected file. | Load the protected workbook, locate the first VbaModule, and try to retrieve its SourceCode via reflection, catching the exception that indicates the macro is locked. | Log the result, confirming that accessing SourceCode throws an exception when the VBA project is password‑protected.
// Common Searches: Aspose.Cells example for password protecting VBA project in C# | how to test VBA macro protection after saving with Aspose.Cells | exception thrown when reading protected VBA module source code Aspose.Cells | using reflection to get VbaModule SourceCode property in protected workbook | C# verify that VBA project is locked in an XLSM file
// Tags: Aspose.Cells VBA project password protection | C# read protected VBA module source code | reflection access VbaModule SourceCode Aspose.Cells | load and save XLSM workbook with macro lock | handle VBA protection exception Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The sample loads an existing .xlsm workbook, applies password protection to its VBA project with Aspose.Cells, saves the file, then reloads it and attempts to read the SourceCode of the first VBA module via reflection. The expected exception is caught and logged, confirming that the VBA project protection is enforced.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";
        const string outputPath = "protected.xlsm";
        const string password = "myPassword";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load an existing workbook that contains a VBA project
            Workbook workbook = new Workbook(inputPath);

            // Verify that the workbook actually contains a VBA project
            if (workbook.VbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Protect the VBA project with a password (read‑only flag set to true)
            workbook.VbaProject.Protect(true, password);

            // Save the workbook after protection
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}' with VBA protection.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error during workbook processing: " + ex.Message);
            return;
        }

        // Attempt to read VBA code to confirm that protection is enforced
        try
        {
            // Reload the protected workbook
            Workbook protectedWb = new Workbook(outputPath);

            // Access the first VBA module (if any)
            if (protectedWb.VbaProject != null && protectedWb.VbaProject.Modules.Count > 0)
            {
                VbaModule vbaModule = protectedWb.VbaProject.Modules[0];

                // Use reflection to get the SourceCode property (avoids compile‑time dependency)
                PropertyInfo sourceProp = typeof(VbaModule).GetProperty("SourceCode",
                    BindingFlags.Public | BindingFlags.Instance);

                if (sourceProp != null)
                {
                    // This line should throw an exception because the project is protected
                    string code = sourceProp.GetValue(vbaModule) as string;

                    // If no exception, protection is not enforced (unexpected)
                    Console.WriteLine("VBA code was read (unexpected):");
                    Console.WriteLine(code);
                }
                else
                {
                    Console.WriteLine("SourceCode property not available in this Aspose.Cells version.");
                }
            }
            else
            {
                Console.WriteLine("No VBA modules found in the protected workbook.");
            }
        }
        catch (Exception ex)
        {
            // Expected outcome: an exception indicating the VBA project is protected
            Console.WriteLine("Failed to read VBA code as expected: " + ex.Message);
        }
    }
}
