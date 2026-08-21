// Title: Digitally Sign a VBA Project in an .xlsm Workbook Using a Windows Store Certificate Thumbprint (Aspose.Cells for .NET)
// Description: This example demonstrates how to load a macro‑enabled Excel workbook, retrieve an X509Certificate2 from the CurrentUser\My Windows certificate store by its thumbprint, create an Aspose.Cells DigitalSignature, sign the workbook's VbaProject, verify the signature status, and save the signed file.
// Keywords: Aspose.Cells | C# | VBA project signing | Excel macro digital signature | Windows certificate store | thumbprint lookup | X509Certificate2 | DigitalSignature API | macro‑enabled workbook | programmatic code signing
// Common Searches: sign VBA project Aspose.Cells C# | retrieve certificate by thumbprint Windows store | digitally sign .xlsm macros programmatically | Aspose.Cells example for VBA digital signature | how to use Windows certificate store with Aspose.Cells
// Developer Intent: The developer needs to apply a digital signature to the VBA project of a macro‑enabled Excel file using a certificate stored in the Windows certificate store, identified by its thumbprint.
// Use Cases: Apply a corporate code‑signing certificate to VBA macros before distribution to ensure authenticity and integrity. | Automate signing of multiple .xlsm files in a CI/CD pipeline by fetching the certificate via thumbprint. | Validate that a workbook’s VBA project is correctly signed and trusted after the signing operation.
// AI Prompts: Generate C# code that signs a VBA project in an .xlsm workbook using a certificate retrieved from the CurrentUser\My store by thumbprint with Aspose.Cells. | Add comprehensive error handling to the VBA signing sample for missing certificates, inaccessible stores, and invalid signatures. | Explain how to programmatically verify the digital signature of a VBA project after signing it with Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigningDemo
{
    // This example demonstrates how to load a macro‑enabled Excel workbook, retrieve an X509Certificate2 from the CurrentUser\My Windows certificate store by its thumbprint, create an Aspose.Cells DigitalSignature, sign the workbook's VbaProject, verify the signature status, and save the signed file.
    public class SignVbaProjectWithStoreCertificate
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
                    Console.WriteLine($"Input workbook not found: '{workbookPath}'.");
                    return;
                }

                // Thumbprint of the certificate stored in the Windows certificate store
                string certificateThumbprint = "YOUR_CERTIFICATE_THUMBPRINT";

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Retrieve the certificate from the CurrentUser\My store using the thumbprint
                X509Certificate2 certificate = null;
                using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
                {
                    store.Open(OpenFlags.ReadOnly);
                    foreach (var cert in store.Certificates)
                    {
                        if (string.Equals(cert.Thumbprint, certificateThumbprint, StringComparison.OrdinalIgnoreCase))
                        {
                            certificate = cert;
                            break;
                        }
                    }
                }

                if (certificate == null)
                {
                    Console.WriteLine($"Certificate with thumbprint '{certificateThumbprint}' not found in the store.");
                    return;
                }

                // Create a DigitalSignature object using the retrieved certificate
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,
                    "Signed by Aspose.Cells VBA signing demo",
                    DateTime.Now);

                // Sign the VBA project
                vbaProject.Sign(digitalSignature);

                // Verify signing status
                Console.WriteLine($"VBA Project IsSigned: {vbaProject.IsSigned}");
                Console.WriteLine($"VBA Project IsValidSigned: {vbaProject.IsValidSigned}");

                // Save the signed workbook
                string signedWorkbookPath = "SignedWorkbook.xlsm";
                workbook.Save(signedWorkbookPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved with signed VBA project to '{signedWorkbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SignVbaProjectWithStoreCertificate.Run();
        }
    }
}
