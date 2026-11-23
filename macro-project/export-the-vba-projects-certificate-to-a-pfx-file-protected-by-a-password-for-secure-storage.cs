// Title: Export a VBA project's digital certificate to a password‑protected PFX file using Aspose.Cells for .NET
// Common Searches: how to extract a VBA project's certificate from an XLSM workbook with Aspose.Cells in C# | C# code to save a VBA macro certificate as a password protected .pfx file | using reflection to retrieve VBA project signature certificate via Aspose.Cells | export X509Certificate2 from an Aspose.Cells VBA project to a PFX file with a password | store VBA project certificate securely as PFX using .NET
// Tags: VBA certificate export to PFX Aspose.Cells | Aspose.Cells VBA signature extraction | C# generate password protected PFX | reflection based VBA certificate access | secure macro certificate storage .NET

using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The example loads an XLSM workbook with Aspose.Cells, uses reflection to access the VBA project's SignatureInfo and its X509Certificate2, then exports the certificate to a password‑protected PFX byte array and writes it to a .pfx file, handling missing files and errors throughout.
class ExportVbaCertificate
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the VBA project
            Workbook workbook = new Workbook(inputPath);

            // Attempt to retrieve the VBA project's certificate via reflection
            X509Certificate2 certificate = null;
            try
            {
                var vbaProject = workbook.VbaProject;
                if (vbaProject != null)
                {
                    // Look for a property named "SignatureInfo"
                    PropertyInfo signatureInfoProp = vbaProject.GetType().GetProperty("SignatureInfo");
                    if (signatureInfoProp != null)
                    {
                        object signatureInfo = signatureInfoProp.GetValue(vbaProject);
                        if (signatureInfo != null)
                        {
                            // Look for a property named "Certificate" inside SignatureInfo
                            PropertyInfo certProp = signatureInfo.GetType().GetProperty("Certificate");
                            if (certProp != null)
                            {
                                certificate = certProp.GetValue(signatureInfo) as X509Certificate2;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving certificate via reflection: {ex.Message}");
            }

            if (certificate != null)
            {
                const string pfxFilePath = "vba_certificate.pfx";
                const string password = "StrongPassword123";

                try
                {
                    // Export the certificate to a byte array in PFX format protected by the password
                    byte[] pfxBytes = certificate.Export(X509ContentType.Pfx, password);

                    // Ensure the directory for the output file exists
                    string outputDir = Path.GetDirectoryName(pfxFilePath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Write the PFX data to the file system
                    File.WriteAllBytes(pfxFilePath, pfxBytes);
                    Console.WriteLine($"Certificate exported successfully to '{pfxFilePath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export certificate: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No certificate is associated with the VBA project or the API does not expose it.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
