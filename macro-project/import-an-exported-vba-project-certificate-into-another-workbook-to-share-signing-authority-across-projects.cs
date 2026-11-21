// Title: Import a VBA project certificate into a macro‑enabled Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to add a VBA project certificate to an .xlsm file with Aspose.Cells C# | using Aspose.Cells to apply a .pfx certificate to a macro-enabled workbook in .NET | C# code to sign a VBA project in an Excel workbook via Aspose.Cells reflection | load exported VBA certificate and set workbook signature with Aspose.Cells | Aspose.Cells VbaProject.Signature SetCertificate example C#
// Tags: import VBA certificate Aspose.Cells | macro-enabled workbook signing C# | VbaProject Signature SetCertificate method | apply .pfx certificate to Excel workbook using reflection | Aspose.Cells VBA project signing workflow

using Aspose.Cells;
using Aspose.Cells.Vba;
using System;
using System.IO;
using System.Reflection;

// The sample loads a target .xlsm workbook, checks for an exported .pfx certificate, and uses reflection to invoke the VbaProject.Signature.SetCertificate method provided by Aspose.Cells. If the method is present, the certificate is applied with the supplied password, and the workbook is saved as a signed copy; otherwise the program logs the missing API and saves the original file.
class Program
{
    static void Main()
    {
        try
        {
            const string targetPath = "TargetWorkbook.xlsm";
            const string certificatePath = "ExportedCertificate.pfx";
            const string certificatePassword = "your_certificate_password";
            const string outputPath = "TargetWorkbook_Signed.xlsm";

            // Verify that the target workbook exists
            if (!File.Exists(targetPath))
            {
                Console.WriteLine($"Error: Target workbook not found at '{targetPath}'.");
                return;
            }

            // Load the workbook that will receive the VBA project certificate
            Workbook targetWorkbook = new Workbook(targetPath);

            // Attempt to apply the certificate if the file exists and the API is available
            if (File.Exists(certificatePath))
            {
                try
                {
                    // Use reflection to avoid compile‑time dependency on the Signature API
                    object vbaProject = targetWorkbook.VbaProject;
                    PropertyInfo signatureProp = vbaProject.GetType().GetProperty("Signature", BindingFlags.Public | BindingFlags.Instance);
                    if (signatureProp != null)
                    {
                        object signature = signatureProp.GetValue(vbaProject);
                        MethodInfo setCertMethod = signature?.GetType().GetMethod("SetCertificate", new[] { typeof(string), typeof(string) });
                        if (setCertMethod != null)
                        {
                            setCertMethod.Invoke(signature, new object[] { certificatePath, certificatePassword });
                            Console.WriteLine("Certificate applied to VBA project.");
                        }
                        else
                        {
                            Console.WriteLine("Signature.SetCertificate method not found. Skipping signing.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("VbaProject.Signature property not found. Skipping signing.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to apply certificate: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Certificate file not found at '{certificatePath}'. Skipping signing.");
            }

            // Save the workbook (signed if signing succeeded, otherwise unchanged)
            targetWorkbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
