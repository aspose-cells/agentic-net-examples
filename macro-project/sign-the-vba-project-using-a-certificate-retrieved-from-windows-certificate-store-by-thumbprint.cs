// Title: C# – Sign a VBA Project in an XLSM Workbook Using a Windows Certificate Store Thumbprint with Aspose.Cells
// Description: Load a macro‑enabled workbook, fetch a personal X509Certificate2 by thumbprint from the Current User store, create an Aspose.Cells DigitalSignature, sign the VBA project, save the file, and verify the signature.
// Keywords: Aspose.Cells | C# | VBA project signing | XLSM digital signature | Windows certificate store | thumbprint lookup | X509Certificate2 | macro‑enabled workbook | programmatic signing | digital signature verification
// Common Searches: sign VBA project Aspose.Cells C# | retrieve certificate by thumbprint .NET | add digital signature to XLSM file | how to sign macro‑enabled workbook programmatically | verify VBA project signature after saving
// Developer Intent: Apply a digital signature to the VBA project of a macro‑enabled workbook using a certificate stored in the Windows certificate store.
// Use Cases: Secure distribution of macro‑enabled workbooks by signing the VBA code with a corporate certificate. | Automate signing of multiple XLSM files in a CI/CD pipeline using a single thumbprint‑identified certificate. | Validate that a signed VBA project remains intact after saving and reloading the workbook.
// AI Prompts: Generate C# code that uses Aspose.Cells to sign the VBA project of an XLSM workbook with a certificate retrieved from the CurrentUser Personal store by thumbprint, then confirm the signature status. | Explain how to locate an X509Certificate2 in the Windows certificate store using a thumbprint and use it to create a DigitalSignature for a VBA project with Aspose.Cells. | Write a C# routine that iterates over a list of macro‑enabled workbooks, signs each VBA project with the same certificate, and logs success or failure for each file.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigning
{
    // Load a macro‑enabled workbook, fetch a personal X509Certificate2 by thumbprint from the Current User store, create an Aspose.Cells DigitalSignature, sign the VBA project, save the file, and verify the signature.
    public class VbaProjectSigner
    {
        public static void Run()
        {
            try
            {
                // Path to the macro-enabled workbook that contains a VBA project
                string workbookPath = "InputWorkbook.xlsm";

                // Verify the input workbook exists
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
                string certificateThumbprint = "YOUR_CERTIFICATE_THUMBPRINT";

                // Retrieve the certificate from the Current User's Personal store
                X509Certificate2 certificate = GetCertificateByThumbprint(certificateThumbprint);
                if (certificate == null)
                {
                    Console.WriteLine("Certificate with the specified thumbprint was not found.");
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

                // Verify the signature
                if (File.Exists(signedWorkbookPath))
                {
                    Workbook verifyWorkbook = new Workbook(signedWorkbookPath);
                    Console.WriteLine("VBA Project IsSigned: " + verifyWorkbook.VbaProject.IsSigned);
                    Console.WriteLine("VBA Project IsValidSigned: " + verifyWorkbook.VbaProject.IsValidSigned);
                }
                else
                {
                    Console.WriteLine($"Failed to save signed workbook: {signedWorkbookPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during VBA project signing:");
                Console.WriteLine(ex.Message);
            }
        }

        private static X509Certificate2 GetCertificateByThumbprint(string thumbprint)
        {
            // Open the personal (My) store of the current user
            using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);

                // Find certificates matching the thumbprint (case‑insensitive, ignore spaces)
                X509Certificate2Collection found = store.Certificates.Find(
                    X509FindType.FindByThumbprint,
                    thumbprint.Replace(" ", string.Empty),
                    false);

                // Return the first matching certificate, or null if none found
                return found.Count > 0 ? found[0] : null;
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
