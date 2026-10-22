// Title: Import an exported VBA project certificate into an XLSM workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells import VBA project certificate into existing XLSM file | load VBA project signature from .cer file using Aspose.Cells .NET | use reflection to call ImportCertificate on VBA project in Aspose.Cells | add signing authority to macro-enabled workbook programmatically with Aspose.Cells | how to apply exported VBA certificate to another workbook in C#
// Tags: import VBA certificate Aspose.Cells | VBA project signature Aspose.Cells | macro-enabled workbook certificate C# | reflection call ImportCertificate Aspose.Cells | share VBA signing authority across XLSM files

using Aspose.Cells;
using System;
using System.IO;

// The example loads a target XLSM workbook, checks for a VBA project, and uses reflection to invoke the Signature.ImportCertificate method (when available) to import an exported .cer certificate, then saves the workbook preserving the macro and the new signing authority.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to the workbook and the exported VBA certificate
            string targetWorkbookPath = "TargetWorkbook.xlsm";
            string exportedCertificatePath = "ExportedVbaCertificate.cer";

            // Verify that the workbook file exists
            if (!File.Exists(targetWorkbookPath))
            {
                Console.WriteLine($"Workbook file not found: {targetWorkbookPath}");
                return;
            }

            // Verify that the certificate file exists
            if (!File.Exists(exportedCertificatePath))
            {
                Console.WriteLine($"Certificate file not found: {exportedCertificatePath}");
                return;
            }

            // Load the target workbook
            Workbook targetWorkbook = new Workbook(targetWorkbookPath);

            // Check if the workbook contains a VBA project
            if (targetWorkbook.VbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project. No certificate will be imported.");
            }
            else
            {
                // Attempt to import the certificate if the API supports it
                try
                {
                    // In newer Aspose.Cells versions the Signature property is available.
                    // Use reflection to call ImportCertificate if the property exists.
                    var signatureProp = targetWorkbook.VbaProject.GetType().GetProperty("Signature");
                    if (signatureProp != null)
                    {
                        var signatureObj = signatureProp.GetValue(targetWorkbook.VbaProject);
                        var importMethod = signatureObj?.GetType().GetMethod("ImportCertificate", new[] { typeof(string) });
                        if (importMethod != null)
                        {
                            importMethod.Invoke(signatureObj, new object[] { exportedCertificatePath });
                            Console.WriteLine("Certificate imported successfully.");
                        }
                        else
                        {
                            Console.WriteLine("ImportCertificate method not found on Signature object.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Signature property not available in this Aspose.Cells version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to import certificate: {ex.Message}");
                }
            }

            // Save the workbook (XLSM preserves VBA)
            string signedWorkbookPath = "TargetWorkbook_Signed.xlsm";
            targetWorkbook.Save(signedWorkbookPath);
            Console.WriteLine($"Workbook saved: {signedWorkbookPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
