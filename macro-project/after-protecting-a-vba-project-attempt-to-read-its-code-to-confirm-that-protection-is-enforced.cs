// Title: C# example: protect an XLSM VBA project with a password using Aspose.Cells and verify enforcement by attempting to read its code
// Common Searches: c# protect vba project in xlsm using aspose.cells | asp.net read vba module source code after workbook protection | how to test if vba project password protection is active with aspose.cells | example of Workbook.VbaProject.Protect and verification in C# | retrieve vba source code from protected workbook using reflection
// Tags: Aspose.Cells VBA project password protection C# | Aspose.Cells VBA module source retrieval via reflection | Workbook.VbaProject.Protect method usage | detect protected VBA project exception handling | verify VBA protection after saving workbook

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Vba; // For VbaProject and VbaModule

namespace AsposeCellsVbaProtectionDemo
{
    // The sample loads an XLSM workbook, applies password protection to its VBA project with Workbook.VbaProject.Protect, saves the file, reloads it, and iterates through each VbaModule. It attempts to obtain the SourceCode property via reflection; successful retrieval indicates protection is not enforced, while an exception confirms that the password protection is active.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsm";
            const string protectedPath = "protected.xlsm";
            const string vbaPassword = "MySecretPassword";

            // Verify the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load an existing workbook that contains a VBA project
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook actually has a VBA project
                if (workbook.VbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Protect the VBA project with a password (first argument: isReadOnly flag)
                workbook.VbaProject.Protect(true, vbaPassword);

                // Save the workbook after protection
                workbook.Save(protectedPath);
                Console.WriteLine($"Workbook saved as '{protectedPath}' with VBA protection.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error during protection/saving:");
                Console.WriteLine(ex.Message);
                return;
            }

            // Attempt to read VBA code after protection
            try
            {
                // Verify the protected workbook exists
                if (!File.Exists(protectedPath))
                {
                    Console.WriteLine($"Protected file '{protectedPath}' not found.");
                    return;
                }

                // Reload the protected workbook to simulate a fresh read
                Workbook protectedWorkbook = new Workbook(protectedPath);

                // Access the VBA project
                var vbaProject = protectedWorkbook.VbaProject;
                if (vbaProject == null)
                {
                    Console.WriteLine("Protected workbook does not contain a VBA project.");
                    return;
                }

                // Iterate through all VBA modules
                foreach (VbaModule module in vbaProject.Modules)
                {
                    // Retrieve source code via reflection (avoids compile‑time dependency on the property)
                    string code = null;
                    PropertyInfo srcProp = typeof(VbaModule).GetProperty("SourceCode", BindingFlags.Public | BindingFlags.Instance);
                    if (srcProp != null)
                    {
                        code = srcProp.GetValue(module) as string;
                    }
                    else
                    {
                        code = "[SourceCode property not available in this Aspose.Cells version]";
                    }

                    Console.WriteLine($"Module: {module.Name}");
                    Console.WriteLine(code);
                }

                // If we reach this point without exception, the protection may not be enforced
                Console.WriteLine("VBA code was read successfully. Protection may not be active.");
            }
            catch (Exception ex)
            {
                // Expected path: an exception is thrown because the VBA project is protected
                Console.WriteLine("Failed to read VBA code as expected. Protection is enforced.");
                Console.WriteLine($"Exception message: {ex.Message}");
            }
        }
    }
}
