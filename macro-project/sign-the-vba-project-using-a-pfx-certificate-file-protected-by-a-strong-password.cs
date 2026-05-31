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
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Paths for input workbook, output workbook and certificate
                string inputPath = "MacroWorkbook.xlsm";
                string outputPath = "MacroWorkbook_Signed.xlsm";
                string certificatePath = "MyCertificate.pfx";
                string certificatePassword = "StrongPassword123!";

                // Verify that required files exist
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input workbook not found: {inputPath}");
                    return;
                }

                if (!File.Exists(certificatePath))
                {
                    Console.WriteLine($"Certificate file not found: {certificatePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Load the certificate (includes private key)
                X509Certificate2 certificate = new X509Certificate2(
                    File.ReadAllBytes(certificatePath), certificatePassword);

                // Create a digital signature instance
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,               // certificate with private key
                    "Signed by Aspose.Cells", // optional comment
                    DateTime.Now);            // signing time

                // Sign the VBA project
                vbaProject.Sign(digitalSignature);

                // Save the signed workbook
                workbook.Save(outputPath, SaveFormat.Xlsm);

                // Verify the signature
                Workbook verifyWorkbook = new Workbook(outputPath);
                Console.WriteLine("VBA Project Signed: " + verifyWorkbook.VbaProject.IsSigned);
                Console.WriteLine("VBA Signature Valid: " + verifyWorkbook.VbaProject.IsValidSigned);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}