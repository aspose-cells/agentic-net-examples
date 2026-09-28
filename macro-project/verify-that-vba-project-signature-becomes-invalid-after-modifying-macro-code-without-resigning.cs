// Title: Verify that editing a signed VBA module invalidates the VBA project signature in an .xlsm workbook using Aspose.Cells for .NET
// AI Prompts: Load a signed .xlsm workbook with Aspose.Cells, read the VbaProject.IsSigned flag, append a comment to a VBA module, re‑evaluate IsSigned, and save the workbook. | Programmatically change the code of the first VBA module in a signed Excel macro file without re‑signing, then confirm that VbaProject.IsSigned returns false. | Show how to detect a broken VBA digital signature after modifying macro source by using the Aspose.Cells VbaProject API in C#.
// Common Searches: aspnet check VBA project signature after editing macro code | c# Aspose.Cells verify signed macro integrity | how to detect invalid VBA digital signature in modified xlsm workbook | Aspose.Cells VbaProject.IsSigned returns false after code change | test if modifying signed VBA module breaks signature using .NET
// Tags: Aspose.Cells VBA signature verification | edit signed VBA module in .xlsm | detect broken VBA digital signature C# | check macro integrity after modification | Aspose.Cells VbaProject API usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads a signed .xlsm workbook, reads the VbaProject.IsSigned property, appends a comment to a VBA module without re‑signing, checks the IsSigned flag again, and saves the file, demonstrating that the signature becomes invalid after the code change.
class VbaSignatureVerification
{
    static void Main()
    {
        try
        {
            const string inputPath = "SignedMacro.xlsm";
            const string outputPath = "ModifiedMacro.xlsm";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains a signed VBA project.
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project embedded in the workbook.
            VbaProject vbaProject = workbook.VbaProject;

            // Check the initial signature status using IsSigned.
            bool isSignedBefore = vbaProject.IsSigned;
            Console.WriteLine($"Workbook is signed before modification: {isSignedBefore}");

            // Modify the VBA code without re‑signing.
            // Append a simple comment to the first module (adjust the name as needed).
            VbaModule module = null;
            try
            {
                module = vbaProject.Modules["Module1"];
            }
            catch (Exception)
            {
                Console.WriteLine("Module 'Module1' not found. Attempting to use the first available module.");
                if (vbaProject.Modules.Count > 0)
                {
                    module = vbaProject.Modules[0];
                }
            }

            if (module != null)
            {
                string originalCode = module.Codes;
                module.Codes = originalCode + "\n' Added comment without re‑signing";
            }
            else
            {
                Console.WriteLine("No VBA module available to modify.");
            }

            // After modification, re‑check the signature status.
            bool isSignedAfter = vbaProject.IsSigned;
            Console.WriteLine($"Workbook is signed after modification: {isSignedAfter}");

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Modified workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
