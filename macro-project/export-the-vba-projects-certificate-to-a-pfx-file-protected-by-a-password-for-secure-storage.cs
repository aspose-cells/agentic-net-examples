// Title: Export a VBA project's digital certificate to a password‑protected PFX file with Aspose.Cells for .NET (C#)
// Common Searches: how to export a VBA project's certificate to a .pfx file using Aspose.Cells in C# | C# code example for saving Excel macro certificate as password protected PFX with Aspose.Cells | retrieve VBA signature from an XLSM workbook and export to PFX using Aspose.Cells | using reflection to call ExportCertificate on VBA project in Aspose.Cells .NET | Aspose.Cells VbaProject ExportCertificate method usage for secure certificate storage
// Tags: export vba certificate to pfx Aspose.Cells | Aspose.Cells VbaProject ExportCertificate API | C# extract VBA macro certificate password protected pfx | reflection based VBA signature export .NET | secure storage of Excel VBA project certificate

using Aspose.Cells;
using Aspose.Cells.Vba;
using System;
using System.IO;
using System.Reflection;

// The sample loads an XLSM workbook, checks for a VBA project, uses reflection to access the project's Signature property, and invokes the ExportCertificate method to write the VBA certificate to a password‑protected .pfx file, handling missing features and runtime errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";

            // Ensure the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the VBA project
            Workbook workbook = new Workbook(inputPath);

            // Get the VBA project from the workbook
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Use reflection to access the Signature property (may not exist in older versions)
            PropertyInfo signatureProp = vbaProject.GetType().GetProperty("Signature", BindingFlags.Public | BindingFlags.Instance);
            if (signatureProp == null)
            {
                Console.WriteLine("VBA signature functionality is not available in the current Aspose.Cells version.");
                return;
            }

            object signature = signatureProp.GetValue(vbaProject);
            if (signature == null)
            {
                Console.WriteLine("No VBA certificate found in the workbook.");
                return;
            }

            // Define the output .pfx file path and the password to protect it
            string pfxPath = "vba_certificate.pfx";
            string password = "StrongPassword123";

            try
            {
                // Use reflection to call ExportCertificate(string, string)
                MethodInfo exportMethod = signature.GetType().GetMethod("ExportCertificate", new[] { typeof(string), typeof(string) });
                if (exportMethod == null)
                {
                    Console.WriteLine("ExportCertificate method is not available on the VBA signature object.");
                    return;
                }

                exportMethod.Invoke(signature, new object[] { pfxPath, password });
                Console.WriteLine($"Certificate exported successfully to {pfxPath}");
            }
            catch (TargetInvocationException tie) when (tie.InnerException != null)
            {
                Console.WriteLine($"Failed to export certificate: {tie.InnerException.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to export certificate: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
