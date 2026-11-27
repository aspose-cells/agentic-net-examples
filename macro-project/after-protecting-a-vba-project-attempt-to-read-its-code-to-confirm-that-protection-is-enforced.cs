// Title: Protect a VBA project in an XLSM workbook and verify protection by attempting to read module code using Aspose.Cells for .NET (C#)
// Common Searches: how to protect a VBA project in an XLSM file with Aspose.Cells for .NET | C# example to test VBA project protection after saving workbook using Aspose.Cells | read VBA module source from a password‑protected workbook with Aspose.Cells C# | verify that VBA code is inaccessible after applying protection via Aspose.Cells | Aspose.Cells VBA project protection and code‑retrieval validation in C#
// Tags: Aspose.Cells protect VBA project C# | retrieve VBA module code Aspose.Cells | programmatic VBA protection verification | XLSM workbook VBA security Aspose.Cells | VbaProject GetCode reflection fallback

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The sample loads an XLSM workbook, applies password‑protected read‑only protection to its VbaProject using Aspose.Cells, saves the file, reloads it, and iterates through each VbaModule attempting to obtain the code via GetCode or the Code property. Exceptions are caught to confirm that protected VBA code cannot be accessed, thereby validating the protection.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook that contains a VBA project.
            string inputPath = "SampleWithVba.xlsm";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project.
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("No VBA project found in the input workbook.");
                return;
            }

            // Protect the VBA project with a password (read‑only flag set to true).
            string password = "MySecretPassword";
            vbaProject.Protect(true, password);

            // Save the workbook with the protected VBA project.
            string protectedPath = "SampleWithVba_Protected.xlsm";
            workbook.Save(protectedPath, SaveFormat.Xlsm);

            // Verify that protection is enforced.
            try
            {
                if (!File.Exists(protectedPath))
                {
                    Console.WriteLine($"Protected file not found: {protectedPath}");
                    return;
                }

                // Reload the protected workbook.
                Workbook protectedWorkbook = new Workbook(protectedPath);
                VbaProject protectedVba = protectedWorkbook.VbaProject;
                if (protectedVba == null)
                {
                    Console.WriteLine("No VBA project found in the protected workbook.");
                    return;
                }

                // Iterate through all VBA modules and attempt to read their code.
                foreach (VbaModule module in protectedVba.Modules)
                {
                    try
                    {
                        // The Code property may not be available in older versions;
                        // use GetCode() if present, otherwise fall back to reflection.
                        string code = null;

                        // Prefer the GetCode() method if it exists.
                        var getCodeMethod = typeof(VbaModule).GetMethod("GetCode");
                        if (getCodeMethod != null)
                        {
                            code = (string)getCodeMethod.Invoke(module, null);
                        }
                        else
                        {
                            // Fallback to the Code property via reflection.
                            var codeProp = typeof(VbaModule).GetProperty("Code");
                            if (codeProp != null)
                            {
                                code = (string)codeProp.GetValue(module);
                            }
                        }

                        Console.WriteLine($"Module: {module.Name}");
                        Console.WriteLine(code ?? "[Unable to retrieve code]");
                    }
                    catch (Exception innerEx)
                    {
                        // Expected when the project is protected.
                        Console.WriteLine($"Failed to read code from module '{module.Name}' as expected.");
                        Console.WriteLine($"Exception: {innerEx.Message}");
                    }
                }

                // If we reach this point without any exception, protection may not be effective.
                Console.WriteLine("VBA code was read without exception – protection may not be applied.");
            }
            catch (Exception ex)
            {
                // Any unexpected error while verifying protection.
                Console.WriteLine("An error occurred while verifying VBA protection.");
                Console.WriteLine($"Exception message: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // General error handling.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
