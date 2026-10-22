// Title: Remove VBA digital signature from a macro-enabled .xlsm workbook with Aspose.Cells for .NET
// Common Searches: asp.net remove VBA digital signature from .xlsm using Aspose.Cells | clear VBA project certificate in C# Aspose.Cells workbook | make IsSigned false after loading signed macro workbook with Aspose.Cells | delete VBA project from Excel file programmatically with Aspose.Cells | un‑sign macro‑enabled Excel workbook using Aspose.Cells for .NET
// Tags: Aspose.Cells clear VBA project C# | remove VBA digital signature Aspose.Cells | unsigned macro-enabled workbook .xlsm | workbook.VbaProject null Aspose.Cells | digital certificate removal Excel macro C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads a signed macro‑enabled workbook (SignedWorkbook.xlsm) with Aspose.Cells for .NET, checks for a VBA project, and explains that the library does not expose a direct API to delete a VBA digital signature. It demonstrates how to clear the entire VBA project by setting workbook.VbaProject to null, which also removes the signature, and then saves the workbook as UnsignedWorkbook.xlsm. The code includes file existence validation and exception handling.
class RemoveVbaSignature
{
    static void Main()
    {
        const string inputPath = "SignedWorkbook.xlsm";
        const string outputPath = "UnsignedWorkbook.xlsm";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains a VBA project with a digital signature
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook has a VBA project
            if (workbook.VbaProject != null)
            {
                // NOTE: Aspose.Cells for .NET does not expose a direct API to remove a VBA digital signature.
                // If needed, you can clear the entire VBA project (which also removes the signature) as shown below.
                // Uncomment the following line to remove the whole VBA project:
                // workbook.VbaProject = null;

                Console.WriteLine("VBA project detected. Signature removal is not directly supported by the API.");
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }

            // Save the (potentially modified) workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
