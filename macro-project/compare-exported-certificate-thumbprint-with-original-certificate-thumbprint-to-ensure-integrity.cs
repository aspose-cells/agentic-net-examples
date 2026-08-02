// Title: Check Excel Workbook Digital Signature Thumbprint Against Original X509 Certificate with Aspose.Cells (C#)
// Description: Loads a PFX certificate, extracts its thumbprint, creates and signs an Excel workbook using Aspose.Cells, saves the file, reloads it, retrieves embedded digital signatures, and compares each signature's certificate thumbprint with the original to confirm integrity.
// Keywords: Aspose.Cells digital signature | C# thumbprint verification | X509Certificate2 Excel | compare certificate thumbprint | signed workbook integrity | Excel digital signature validation | Aspose.Cells API thumbprint | certificate thumbprint check
// Common Searches: Aspose.Cells compare signature thumbprint with original certificate | C# verify Excel digital signature thumbprint | How to validate workbook signature using Aspose.Cells | Retrieve thumbprint from signed Excel file Aspose | Check certificate integrity in signed Excel workbook .NET
// Developer Intent: Confirm that the thumbprint of the digital signature embedded in an Excel file matches the thumbprint of the source X509 certificate.
// Use Cases: Automated integrity check for Excel reports signed with a corporate PFX before distribution. | Batch verification of signed financial statements to detect tampering. | Audit logging of mismatched thumbprints, including signer details and timestamp.
// AI Prompts: Write C# code that opens a signed Excel file with Aspose.Cells and throws an exception if any embedded signature thumbprint differs from a given certificate thumbprint. | Create a method that extracts all digital signatures from a workbook and returns true only when every signature's thumbprint equals the supplied X509 thumbprint. | Suggest a logging strategy that records mismatched thumbprint, signer name, and signing time when validating Aspose.Cells digital signatures.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsSignatureDemo
{
    // Loads a PFX certificate, extracts its thumbprint, creates and signs an Excel workbook using Aspose.Cells, saves the file, reloads it, retrieves embedded digital signatures, and compares each signature's certificate thumbprint with the original to confirm integrity.
    public class ThumbprintComparison
    {
        public static void Run()
        {
            try
            {
                // Path to the certificate file (PFX) and its password
                string certificatePath = "mycert.pfx";
                string certificatePassword = "password";

                if (!File.Exists(certificatePath))
                {
                    Console.WriteLine($"Certificate file not found: {certificatePath}");
                    return;
                }

                // Load the original certificate and obtain its thumbprint
                X509Certificate2 originalCertificate = new X509Certificate2(certificatePath, certificatePassword);
                string originalThumbprint = originalCertificate.Thumbprint;
                Console.WriteLine("Original Certificate Thumbprint: " + originalThumbprint);

                // Create a new workbook and add some data
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Document to be signed");

                // Create a digital signature using the original certificate
                DigitalSignature signature = new DigitalSignature(originalCertificate, "Demo Signature", DateTime.Now);
                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                signatures.Add(signature);

                // Add the digital signature to the workbook and save it
                workbook.AddDigitalSignature(signatures);
                string signedFilePath = "SignedWorkbook.xlsx";
                workbook.Save(signedFilePath, SaveFormat.Xlsx);
                Console.WriteLine("Workbook signed and saved to: " + signedFilePath);

                // Load the signed workbook
                if (!File.Exists(signedFilePath))
                {
                    Console.WriteLine($"Signed workbook not found: {signedFilePath}");
                    return;
                }

                Workbook signedWorkbook = new Workbook(signedFilePath);

                // Retrieve the digital signatures from the signed workbook
                DigitalSignatureCollection loadedSignatures = signedWorkbook.GetDigitalSignature();

                // Compare thumbprints of each loaded signature with the original thumbprint
                foreach (DigitalSignature loadedSignature in loadedSignatures)
                {
                    string loadedThumbprint = loadedSignature.Certificate.Thumbprint;
                    bool thumbprintsMatch = string.Equals(originalThumbprint, loadedThumbprint, StringComparison.OrdinalIgnoreCase);

                    Console.WriteLine("Loaded Signature Thumbprint: " + loadedThumbprint);
                    Console.WriteLine("Thumbprints match: " + thumbprintsMatch);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ThumbprintComparison.Run();
        }
    }
}
