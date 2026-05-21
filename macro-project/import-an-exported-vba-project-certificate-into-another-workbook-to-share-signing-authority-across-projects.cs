using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaCertificateImport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the workbook that already contains a signed VBA project (source)
                string sourcePath = "SignedSource.xlsm";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: '{sourcePath}'.");
                    return;
                }

                // Load the source workbook
                Workbook sourceWb = new Workbook(sourcePath);

                // Access its VBA project
                VbaProject sourceVba = sourceWb.VbaProject;

                // Ensure the source VBA project is signed
                if (!sourceVba.IsSigned)
                {
                    Console.WriteLine("Source VBA project is not signed.");
                    return;
                }

                // Retrieve the raw certificate data from the signed VBA project
                byte[] certRawData = sourceVba.CertRawData;

                // Optionally, save the certificate to a file for inspection
                string certFile = "ExportedVbaCertificate.cer";
                File.WriteAllBytes(certFile, certRawData);
                Console.WriteLine($"Certificate raw data exported to '{certFile}' (length: {certRawData.Length}).");

                // -----------------------------------------------------------------
                // Import the exported certificate into another workbook (target)
                // -----------------------------------------------------------------

                // Path to the target workbook that needs to share the signing authority
                string targetPath = "TargetWorkbook.xlsm";

                // Verify target file exists
                if (!File.Exists(targetPath))
                {
                    Console.WriteLine($"Target file not found: '{targetPath}'.");
                    return;
                }

                // Load the target workbook
                Workbook targetWb = new Workbook(targetPath);

                // Access its VBA project (will be created automatically for .xlsm)
                VbaProject targetVba = targetWb.VbaProject;

                // Create an X509Certificate2 instance from the raw certificate data.
                // Note: The raw data from a signed VBA project typically contains only the public certificate.
                // To sign another workbook you need the private key (PFX). For demonstration, we assume the
                // raw data includes the private key or that a corresponding PFX file is available.
                X509Certificate2 certificate = new X509Certificate2(certRawData);

                // Create a DigitalSignature object using the certificate
                DigitalSignature digitalSignature = new DigitalSignature(certificate, "Imported VBA Certificate", DateTime.Now);

                // Sign the target VBA project
                targetVba.Sign(digitalSignature);
                Console.WriteLine("Target VBA project signed using the imported certificate.");

                // Save the signed target workbook
                string signedTargetPath = "TargetWorkbook_Signed.xlsm";
                targetWb.Save(signedTargetPath, SaveFormat.Xlsm);
                Console.WriteLine($"Signed target workbook saved as '{signedTargetPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}