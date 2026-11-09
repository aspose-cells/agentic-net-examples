// Title: Remove a VBA digital signature from an XLSM workbook using Aspose.Cells for .NET
// Common Searches: C# remove digital signature from VBA project in an XLSM file with Aspose.Cells | Aspose.Cells clear VBA macro certificate and verify IsSigned property | How to delete VBA project signature programmatically in .NET | Set VbaProject DigitalSignature to null using reflection in C# | Verify VBA project is unsigned after removing signature with Aspose.Cells
// Tags: Aspose.Cells VBA signature removal | C# clear VBA project digital certificate | reflection set VbaProject DigitalSignature null | evaluate VbaProject IsSigned flag | export unsigned XLSM workbook

using Aspose.Cells;
using System;
using System.IO;

// The example loads a signed XLSM workbook, checks the VbaProject.IsSigned flag, attempts to remove the VBA digital signature either via the RemoveDigitalSignature method or by clearing the DigitalSignature property through reflection, re‑evaluates the IsSigned status, and saves the workbook without the signature.
class Program
{
    static void Main()
    {
        const string inputPath = "SignedVba.xlsm";
        const string outputPath = "UnsignedVba.xlsm";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains a signed VBA project
            Workbook workbook = new Workbook(inputPath);

            // Attempt to remove the digital signature if the VBA project is signed
            if (workbook.VbaProject != null && workbook.VbaProject.IsSigned)
            {
                try
                {
                    // The RemoveDigitalSignature method is not available in older versions of Aspose.Cells.
                    // As an alternative, clear the digital signature by setting it to null if the property exists.
                    // If the API is unavailable, this block will simply skip the removal.
                    // Uncomment the line below if your version supports it:
                    // workbook.VbaProject.RemoveDigitalSignature();

                    // Fallback: try to clear the signature via reflection (safe no‑op if not present)
                    var prop = workbook.VbaProject.GetType().GetProperty("DigitalSignature");
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(workbook.VbaProject, null);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to remove digital signature: {ex.Message}");
                }
            }

            // Verify that the VBA project is no longer signed
            bool isSigned = workbook.VbaProject != null && workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project signed: " + isSigned);

            // Save the workbook (signature removed if possible)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
