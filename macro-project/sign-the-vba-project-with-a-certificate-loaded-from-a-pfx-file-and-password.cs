using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigningDemo
{
    public class VbaProjectSigner
    {
        public static void Run()
        {
            try
            {
                // Path to the macro-enabled workbook (XLSM) that contains a VBA project
                string workbookPath = "InputWorkbook.xlsm";

                // Verify workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Ensure the workbook actually contains a VBA project
                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Path to the signing certificate (PFX) and its password
                string certPath = "MyCertificate.pfx";
                string certPassword = "pfxPassword";

                // Verify certificate file exists
                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                // Load the signing certificate
                X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);

                // Create a DigitalSignature instance using the certificate
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,               // certificate with private key
                    "Signed by Aspose.Cells", // optional comment
                    DateTime.Now);            // signing time

                // Sign the VBA project
                vbaProject.Sign(digitalSignature);

                // Save the signed workbook as a macro-enabled file
                string signedWorkbookPath = "SignedWorkbook.xlsm";
                workbook.Save(signedWorkbookPath, SaveFormat.Xlsm);

                // Optional: Verify the signature after saving
                Workbook verifyWorkbook = new Workbook(signedWorkbookPath);
                Console.WriteLine("VBA Project Signed: " + verifyWorkbook.VbaProject.IsSigned);
                Console.WriteLine("Signature Valid: " + verifyWorkbook.VbaProject.IsValidSigned);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            VbaProjectSigner.Run();
        }
    }
}