using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class ExportVbaCertificate
    {
        public static void Run()
        {
            try
            {
                // Path to the workbook that contains a signed VBA project
                string workbookPath = "SignedWorkbook.xlsm";

                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file '{workbookPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Verify that the VBA project is signed and certificate data is available
                if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
                {
                    // Retrieve the raw certificate bytes
                    byte[] certData = vbaProject.CertRawData;

                    // Create an X509Certificate2 instance from the raw data
                    X509Certificate2 certificate = new X509Certificate2(certData);

                    // Define a password to protect the exported PFX file
                    string pfxPassword = "StrongPassword123";

                    // Export the certificate to a PFX (PKCS#12) byte array, protected by the password
                    byte[] pfxBytes = certificate.Export(X509ContentType.Pfx, pfxPassword);

                    // Save the PFX file to disk
                    string pfxPath = "VbaProjectCertificate.pfx";
                    File.WriteAllBytes(pfxPath, pfxBytes);

                    Console.WriteLine($"Certificate exported successfully to '{pfxPath}'.");
                }
                else
                {
                    Console.WriteLine("The VBA project is not signed or does not contain certificate data.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificate.Run();
        }
    }
}