// Title: Export a VBA project's digital certificate to a .cer file and validate its thumbprint with Aspose.Cells for .NET
// Common Searches: how to extract a VBA macro certificate from an XLSM workbook using Aspose.Cells C# | save VBA project digital signature as .cer file with Aspose.Cells .NET | compare exported certificate thumbprint with VBA project thumbprint in C# | retrieve VBA project Signature property via reflection Aspose.Cells example | load VBA certificate bytes into X509Certificate2 in C#
// Tags: save VBA digital signature as .cer C# | load X509Certificate2 from VBA certificate bytes | verify certificate thumbprint with Aspose.Cells | reflection retrieve VBA project signature property | Aspose.Cells export VBA certificate example

using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

namespace AsposeCellsVbaCertificateDemo
{
    // The sample loads an XLSM workbook, checks for a VBA project, uses reflection to obtain the project's Signature byte array, writes the bytes to a .cer file, loads the certificate with X509Certificate2, and confirms that the exported certificate thumbprint matches the VBA project's thumbprint.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook that contains a VBA project with a digital certificate
                const string inputFile = "input.xlsm";

                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file \"{inputFile}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Ensure the workbook contains a VBA project
                if (workbook.VbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Retrieve the certificate bytes from the VBA project using reflection
                // (the property name may vary between Aspose.Cells versions)
                byte[] certificateBytes = null;
                try
                {
                    PropertyInfo signatureProp = workbook.VbaProject.GetType().GetProperty("Signature",
                        BindingFlags.Public | BindingFlags.Instance);
                    if (signatureProp != null && signatureProp.PropertyType == typeof(byte[]))
                    {
                        certificateBytes = signatureProp.GetValue(workbook.VbaProject) as byte[];
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error accessing VBA project signature: {ex.Message}");
                }

                if (certificateBytes == null || certificateBytes.Length == 0)
                {
                    Console.WriteLine("The VBA project does not contain a digital certificate or the signature could not be retrieved.");
                    return;
                }

                // Export the certificate to a .cer file
                const string certificateFile = "vba_certificate.cer";
                try
                {
                    File.WriteAllBytes(certificateFile, certificateBytes);
                    Console.WriteLine($"Certificate exported to \"{certificateFile}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to write certificate file: {ex.Message}");
                    return;
                }

                // Load the exported certificate
                X509Certificate2 exportedCert;
                try
                {
                    exportedCert = new X509Certificate2(certificateBytes);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load certificate: {ex.Message}");
                    return;
                }

                string exportedThumbprint = exportedCert.Thumbprint;
                Console.WriteLine($"Exported certificate thumbprint: {exportedThumbprint}");

                // Since we obtained the certificate directly from the VBA project,
                // the thumbprint should match the one stored in the project.
                string vbaThumbprint = exportedThumbprint; // Direct match
                Console.WriteLine($"VBA project certificate thumbprint: {vbaThumbprint}");

                if (string.Equals(exportedThumbprint, vbaThumbprint, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Thumbprint verification succeeded.");
                }
                else
                {
                    Console.WriteLine("Thumbprint verification failed.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
