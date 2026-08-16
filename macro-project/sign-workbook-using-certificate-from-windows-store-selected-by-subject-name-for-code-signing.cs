// Title: Sign an Excel workbook with a Windows Store code‑signing certificate using Aspose.Cells (C#)
// Description: C# example that opens the current user's personal certificate store, locates a code‑signing X509Certificate2 by subject name, creates an Aspose.Cells DigitalSignature, adds it to a workbook and saves the signed file.
// Keywords: Aspose.Cells digital signature C# | sign Excel workbook Windows certificate store | X509Certificate2 code signing Aspose | C# find certificate by subject name | apply digital signature to .xlsx | Windows Store certificate Aspose.Cells
// Common Searches: How to sign an Excel file with a Windows Store certificate in C# | Aspose.Cells add digital signature using X509Certificate2 | C# retrieve code signing certificate by subject name | Save signed workbook with Aspose.Cells .NET | Digital signature for Excel workbook using Windows certificate store
// Developer Intent: Apply a code‑signing certificate from the Windows certificate store to an Excel workbook with Aspose.Cells.
// Use Cases: Automatically sign generated financial reports before distribution to guarantee authenticity. | Integrate workbook signing into CI/CD pipelines using a corporate certificate stored on build agents. | Enforce compliance by delivering only digitally signed workbooks to external partners.
// AI Prompts: Write C# code that loads an existing workbook, selects a certificate by thumbprint from the Windows store, and signs the file with Aspose.Cells. | Explain how to handle missing or private‑key‑less certificates when signing an Excel workbook using Aspose.Cells. | Provide a step‑by‑step guide to verify a workbook's digital signature after it has been saved with Aspose.Cells.

using System;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    // C# example that opens the current user's personal certificate store, locates a code‑signing X509Certificate2 by subject name, creates an Aspose.Cells DigitalSignature, adds it to a workbook and saves the signed file.
    class Program
    {
        static void Main()
        {
            // Subject name of the code signing certificate (without CN= prefix if you prefer)
            const string certificateSubjectName = "MyCodeSigningCert";

            // Open the current user's personal certificate store
            using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);

                // Find certificates that match the subject name and have a private key
                X509Certificate2 certificate = store.Certificates
                    .Find(X509FindType.FindBySubjectName, certificateSubjectName, validOnly: false)
                    .Cast<X509Certificate2>()
                    .FirstOrDefault(cert => cert.HasPrivateKey);

                if (certificate == null)
                {
                    Console.WriteLine($"Certificate with subject name '{certificateSubjectName}' not found or does not contain a private key.");
                    return;
                }

                // Create a digital signature using the found certificate
                DigitalSignature signature = new DigitalSignature(
                    certificate,
                    "Signed by Aspose.Cells using Windows Store certificate",
                    DateTime.UtcNow);

                // Add the signature to a collection
                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                signatures.Add(signature);

                // Create or load a workbook (here we create a new one)
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Workbook signed with Windows Store certificate");

                // Apply the digital signature to the workbook
                workbook.SetDigitalSignature(signatures);

                // Save the signed workbook
                const string outputPath = "SignedWorkbook.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook signed and saved to '{outputPath}'.");
            }
        }
    }
}
