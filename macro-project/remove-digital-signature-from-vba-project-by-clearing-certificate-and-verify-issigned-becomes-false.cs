// Title: Remove a VBA project's digital signature from a signed .xlsm workbook using Aspose.Cells for .NET
// AI Prompts: Load a signed .xlsm file with Aspose.Cells, detect if its VBA project is signed, invoke the RemoveDigitalSignature method via reflection, and save the workbook without a signature. | Check the IsSigned property of a workbook's VBA project, call the appropriate API to clear the certificate, and confirm the signature status before saving. | Programmatically strip a digital signature from a VBA project in C#, handling cases where the RemoveDigitalSignature method is unavailable.
// Common Searches: C# Aspose.Cells how to delete VBA digital signature from .xlsm file | remove certificate from VBA project using Aspose.Cells .NET API | check and clear VBA IsSigned flag programmatically in C# | Aspose.Cells reflection RemoveDigitalSignature method example | save unsigned macro-enabled workbook after removing signature
// Tags: remove digital signature VBA Aspose.Cells | clear VBA project certificate .NET | invoke RemoveDigitalSignature reflection C# | verify VBA IsSigned property Aspose | save unsigned macro-enabled workbook

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The example loads a signed .xlsm workbook, checks if its VBA project is signed, uses reflection to call RemoveDigitalSignature when available, verifies that IsSigned is false, and saves the workbook as an unsigned file.
class Program
{
    static void Main()
    {
        const string inputPath = "SignedVba.xlsm";
        const string outputPath = "UnsignedVba.xlsm";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing a signed VBA project
            var workbook = new Workbook(inputPath);

            // Access the VBA project
            var vbaProject = workbook.VbaProject;
            if (vbaProject != null)
            {
                // Remove the digital signature if the project is signed
                if (vbaProject.IsSigned)
                {
                    // Attempt to invoke RemoveDigitalSignature via reflection (available in newer versions)
                    MethodInfo removeMethod = vbaProject.GetType().GetMethod("RemoveDigitalSignature",
                        BindingFlags.Instance | BindingFlags.Public);
                    if (removeMethod != null)
                    {
                        removeMethod.Invoke(vbaProject, null);
                        Console.WriteLine("Digital signature removed from VBA project.");
                    }
                    else
                    {
                        Console.WriteLine("RemoveDigitalSignature method not available in this Aspose.Cells version.");
                    }
                }

                // Verify that the VBA project is no longer signed
                Console.WriteLine($"VBA project signed: {vbaProject.IsSigned}");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook; the VBA project will now be unsigned (if removal succeeded)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
