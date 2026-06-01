using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigning
{
    class Program
    {
        static void Main()
        {
            // Path to the macro-enabled workbook that contains a VBA project
            string inputPath = "MacroWorkbook.xlsm";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Thumbprint of the certificate stored in the Windows certificate store
            string certThumbprint = "YOUR_CERTIFICATE_THUMBPRINT".Replace(" ", string.Empty).ToUpperInvariant();

            // Open the personal (My) store of the current user
            using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);

                // Find the certificate by thumbprint
                X509Certificate2Collection found = store.Certificates.Find(
                    X509FindType.FindByThumbprint,
                    certThumbprint,
                    validOnly: false);

                if (found.Count == 0)
                {
                    Console.WriteLine($"Certificate with thumbprint {certThumbprint} not found in the store.");
                    return;
                }

                X509Certificate2 certificate = found[0];

                // Create a digital signature using the certificate
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,
                    "Signed by Aspose.Cells VBA signing routine",
                    DateTime.Now);

                // Sign the VBA project (method rule: Sign)
                vbaProject.Sign(digitalSignature);
            }

            // Save the signed workbook (lifecycle rule: save)
            string outputPath = "SignedMacroWorkbook.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            // Verify signing status
            Workbook verifyWb = new Workbook(outputPath);
            Console.WriteLine("VBA Project IsSigned: " + verifyWb.VbaProject.IsSigned);
            Console.WriteLine("VBA Project IsValidSigned: " + verifyWb.VbaProject.IsValidSigned);
        }
    }
}