// Title: Sign a VBA project in an .xlsm workbook using a Windows certificate store thumbprint – Aspose.Cells for .NET
// Description: Loads a macro‑enabled workbook, fetches an X509Certificate2 from the Current User Personal store by thumbprint, creates a DigitalSignature, signs the workbook's VbaProject, saves the signed file, and validates the signature using Aspose.Cells APIs.
// Keywords: Aspose.Cells VBA signing | C# certificate thumbprint | digital signature Excel macro | sign .xlsm workbook programmatically | Windows certificate store Aspose
// Common Searches: sign VBA project Aspose.Cells C# | retrieve certificate by thumbprint .NET | add digital signature to Excel macro workbook | how to sign .xlsm file with Windows store certificate | verify signed VBA project programmatically
// Developer Intent: Programmatically apply a digital signature to the VBA project of a macro‑enabled Excel workbook using a certificate retrieved from the Windows certificate store.
// Use Cases: Deploy corporate‑signed macros to guarantee authenticity and prevent tampering. | Automate compliance by signing VBA projects during a build or release pipeline. | Validate that signed VBA code remains intact after distribution or storage.
// AI Prompts: Generate C# code that signs a VBA project with a certificate fetched by thumbprint from the CurrentUser store using Aspose.Cells. | Explain how to handle missing private keys when signing a VBA project with a Windows store certificate. | Show how to programmatically verify the validity of a signed VBA project after saving the workbook.

using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigningDemo
{
    // Loads a macro‑enabled workbook, fetches an X509Certificate2 from the Current User Personal store by thumbprint, creates a DigitalSignature, signs the workbook's VbaProject, saves the signed file, and validates the signature using Aspose.Cells APIs.
    public class VbaProjectSignWithStoreCertificate
    {
        public static void Run()
        {
            try
            {
                // Path to the macro-enabled workbook that contains a VBA project
                string workbookPath = "InputWorkbook.xlsm";

                // Ensure the input workbook exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Input workbook not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Thumbprint of the certificate stored in the Windows certificate store
                string thumbprint = "YOUR_CERTIFICATE_THUMBPRINT".Replace(" ", string.Empty).ToUpperInvariant();

                // Retrieve the certificate from the Current User's Personal store
                X509Certificate2 certificate = null;
                using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
                {
                    store.Open(OpenFlags.ReadOnly);
                    var certCollection = store.Certificates.Find(
                        X509FindType.FindByThumbprint,
                        thumbprint,
                        validOnly: false);

                    certificate = certCollection.OfType<X509Certificate2>().FirstOrDefault();
                }

                if (certificate == null)
                {
                    Console.WriteLine($"Certificate with thumbprint {thumbprint} not found in the store.");
                    return;
                }

                if (!certificate.HasPrivateKey)
                {
                    Console.WriteLine("The selected certificate does not contain a private key required for signing.");
                    return;
                }

                // Create a digital signature using the retrieved certificate
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,
                    "Signed VBA Project",
                    DateTime.Now);

                // Sign the VBA project
                vbaProject.Sign(digitalSignature);

                // Save the signed workbook
                string signedWorkbookPath = "SignedWorkbook.xlsm";
                workbook.Save(signedWorkbookPath, SaveFormat.Xlsm);
                Console.WriteLine($"Signed workbook saved to: {signedWorkbookPath}");

                // Verify signing status
                if (File.Exists(signedWorkbookPath))
                {
                    Workbook verifyWorkbook = new Workbook(signedWorkbookPath);
                    Console.WriteLine("VBA Project IsSigned: " + verifyWorkbook.VbaProject.IsSigned);
                    Console.WriteLine("VBA Project IsValidSigned: " + verifyWorkbook.VbaProject.IsValidSigned);
                }
                else
                {
                    Console.WriteLine("Failed to save the signed workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the signing process:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            VbaProjectSignWithStoreCertificate.Run();
        }
    }
}
