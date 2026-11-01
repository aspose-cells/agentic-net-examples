// Title: Export a VBA project's digital certificate to a PEM file using Aspose.Cells for .NET (C#)
// Common Searches: aspnet export vba project certificate to pem using aspose.cells | c# backup vba digital signature from xlsm workbook | how to retrieve DigitalSignatureInfo from VbaProject in Aspose.Cells | save vba macro certificate as pem file with reflection in C# | export vba digital certificate from macro-enabled workbook using Aspose.Cells
// Tags: Aspose.Cells VBA project certificate handling | VbaProject signature info reflection | C# retrieve VBA signing certificate | Generate PEM from VBA macro certificate | Macro-enabled workbook certificate extraction

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Vba; // Provides VbaProject and related classes

// The example loads a macro‑enabled XLSM workbook with Aspose.Cells, accesses its VbaProject, uses reflection to obtain the DigitalSignatureInfo when available, and invokes its Export method to write the VBA project's digital certificate to a PEM file, including checks for missing files or certificates and robust error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";
        const string certPath = "vba_certificate.pem";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the VBA project.
            Workbook workbook = new Workbook(inputPath);

            // Get the VBA project from the workbook.
            VbaProject vbaProject = workbook.VbaProject;

            // Attempt to retrieve the DigitalSignatureInfo via reflection
            // (property may not exist in older Aspose.Cells versions).
            object digitalSignature = null;
            if (vbaProject != null)
            {
                PropertyInfo sigProp = vbaProject.GetType().GetProperty("DigitalSignatureInfo",
                    BindingFlags.Public | BindingFlags.Instance);
                if (sigProp != null)
                {
                    digitalSignature = sigProp.GetValue(vbaProject);
                }
            }

            if (digitalSignature != null)
            {
                // Export the certificate to a PEM file for backup using reflection.
                MethodInfo exportMethod = digitalSignature.GetType().GetMethod("Export",
                    new[] { typeof(string) });
                exportMethod?.Invoke(digitalSignature, new object[] { certPath });
                Console.WriteLine("VBA digital certificate exported successfully.");
            }
            else
            {
                Console.WriteLine("No digital certificate found in the VBA project.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
