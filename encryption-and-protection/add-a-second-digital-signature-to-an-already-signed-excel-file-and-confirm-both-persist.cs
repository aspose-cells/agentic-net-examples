// Title: Add a Second Digital Signature to an Existing Excel Workbook with Aspose.Cells for .NET
// Description: Load a signed .xlsx file, retrieve its DigitalSignatureCollection, create a new DigitalSignature from a second X509Certificate2 (.pfx), append it, save the workbook, then reload and enumerate the signatures to confirm both the original and new signatures are present and valid.
// Keywords: Aspose.Cells | C# digital signature | multiple Excel signatures | add second signature .NET | X509Certificate2 pfx | DigitalSignatureCollection | verify Excel signatures | programmatic signing Excel | Excel workbook security
// Common Searches: add another digital signature to a signed Excel file Aspose.Cells | verify multiple signatures in an .xlsx using C# | append second X509 certificate to Excel workbook | count digital signatures in Excel with Aspose.Cells | preserve existing signatures when adding a new one
// Developer Intent: Programmatically add a second X509‑based digital signature to an already signed Excel workbook and ensure both signatures remain after saving.
// Use Cases: Add an approver’s signature to a contract workbook that already contains a manager’s signature, maintaining a complete audit trail. | Insert a timestamp signature into a financial report that was previously signed, then validate that the file now holds at least two signatures. | Automate multi‑step signing workflows by loading a signed workbook, applying an additional certificate, and confirming the total signature count for compliance.
// AI Prompts: Generate C# code using Aspose.Cells to load a signed .xlsx, add a new digital signature from a .pfx file, and save the file while keeping existing signatures. | Write a method that returns true if an Excel workbook contains two or more digital signatures, using Aspose.Cells in .NET. | Explain error handling for loading X509 certificates and how to validate each signature after adding multiple digital signatures with Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    // Load a signed .xlsx file, retrieve its DigitalSignatureCollection, create a new DigitalSignature from a second X509Certificate2 (.pfx), append it, save the workbook, then reload and enumerate the signatures to confirm both the original and new signatures are present and valid.
    public class AddSecondSignature
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Paths to the original signed workbook and the certificate files
            string signedWorkbookPath = "SignedWorkbook.xlsx";
            string certificatePath1 = "cert1.pfx"; // not used in this demo but kept for reference
            string certificatePath2 = "cert2.pfx";
            string certificatePassword = "1234567890";

            // Verify that the signed workbook exists
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"Workbook file not found: {signedWorkbookPath}");
                return;
            }

            // Verify that the second certificate exists
            if (!File.Exists(certificatePath2))
            {
                Console.WriteLine($"Certificate file not found: {certificatePath2}");
                return;
            }

            // Load the already signed workbook
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Retrieve the existing digital signature collection (may be null if none)
            DigitalSignatureCollection signatureCollection = workbook.GetDigitalSignature();
            if (signatureCollection == null)
            {
                signatureCollection = new DigitalSignatureCollection();
            }

            // Load the second certificate
            X509Certificate2 secondCertificate;
            try
            {
                // Use Import to avoid obsolete constructor warning
                byte[] certData = File.ReadAllBytes(certificatePath2);
                secondCertificate = new X509Certificate2(certData, certificatePassword);
            }
            catch (Exception certEx)
            {
                Console.WriteLine($"Failed to load certificate: {certEx.Message}");
                return;
            }

            // Create a new digital signature using the second certificate
            DigitalSignature secondSignature = new DigitalSignature(
                secondCertificate,
                "Second signature added by Aspose.Cells",
                DateTime.Now);

            // Add the new signature to the collection
            signatureCollection.Add(secondSignature);

            // Apply the updated collection back to the workbook
            workbook.SetDigitalSignature(signatureCollection);

            // Save the workbook with the additional signature
            string outputPath = "SignedWorkbook_WithSecondSignature.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved with second signature: {outputPath}");

            // Verify that both signatures persist by reloading the file and counting signatures
            Workbook verificationWorkbook = new Workbook(outputPath);
            DigitalSignatureCollection verificationCollection = verificationWorkbook.GetDigitalSignature();

            int signatureCount = 0;
            if (verificationCollection != null)
            {
                foreach (DigitalSignature sig in verificationCollection)
                {
                    signatureCount++;
                    Console.WriteLine($"Signature {signatureCount}: Comments = {sig.Comments}, SignTime = {sig.SignTime}, IsValid = {sig.IsValid}");
                }
            }

            Console.WriteLine($"Total digital signatures after adding second one: {signatureCount}");
        }
    }
}
