// Title: Validate Excel Digital Signature Thumbprint with Aspose.Cells for .NET
// Description: Loads a PFX certificate, signs a new workbook using Aspose.Cells, saves the file, reloads it, extracts the embedded digital signatures, retrieves each signature's certificate thumbprint, and compares those thumbprints to the original certificate's thumbprint to confirm signature integrity.
// Keywords: Aspose.Cells | C# digital signature | Excel workbook signing | certificate thumbprint verification | X509Certificate2 | DigitalSignatureCollection | .NET | signature integrity | compare thumbprints | load signed workbook
// Common Searches: How to verify Excel digital signature thumbprint using Aspose.Cells .NET | Compare original certificate thumbprint with signed workbook thumbprint C# | Validate workbook signature integrity Aspose.Cells | Extract certificate from signed Excel file Aspose.Cells | Check if Excel file was signed with a specific certificate
// Developer Intent: Confirm that a signed Excel file was created with the expected certificate by matching thumbprints.
// Use Cases: Post‑signing validation in automated document pipelines | Security audit to detect tampered or re‑signed Excel files | Runtime check before processing a workbook to ensure trusted signing | Batch verification of multiple signed workbooks against a trusted thumbprint
// AI Prompts: Write C# code that loads a signed .xlsx, extracts the digital signature certificate with Aspose.Cells, and compares its thumbprint to a given value. | Explain how to handle missing or expired certificates when verifying Excel signatures using Aspose.Cells. | Provide best practices for storing trusted thumbprints and performing integrity checks on signed workbooks in .NET. | Show how to iterate over multiple signatures in a workbook and report mismatched thumbprints.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsSignatureIntegrityDemo
{
    // Loads a PFX certificate, signs a new workbook using Aspose.Cells, saves the file, reloads it, extracts the embedded digital signatures, retrieves each signature's certificate thumbprint, and compares those thumbprints to the original certificate's thumbprint to confirm signature integrity.
    public class Program
    {
        public static void Main()
        {
            // Path to the certificate (PFX) and its password
            string certPath = "mycert.pfx";
            string certPassword = "password";

            // Load the original certificate (contains private key)
            X509Certificate2 originalCert = new X509Certificate2(certPath, certPassword);
            string originalThumbprint = originalCert.Thumbprint;
            Console.WriteLine("Original Certificate Thumbprint: " + originalThumbprint);

            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Document to be signed");

            // Create a digital signature using the original certificate
            DigitalSignature signature = new DigitalSignature(originalCert, "Signed by Aspose", DateTime.Now);
            DigitalSignatureCollection signatures = new DigitalSignatureCollection();
            signatures.Add(signature);

            // Add the digital signature to the workbook
            workbook.AddDigitalSignature(signatures);

            // Save the signed workbook (lifecycle rule: save)
            string signedPath = "SignedWorkbook.xlsx";
            workbook.Save(signedPath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook signed and saved to: " + signedPath);

            // Load the signed workbook (lifecycle rule: load)
            Workbook loadedWorkbook = new Workbook(signedPath);

            // Retrieve the digital signatures from the loaded workbook
            DigitalSignatureCollection loadedSignatures = loadedWorkbook.GetDigitalSignature();

            // Compare thumbprints of each loaded signature with the original thumbprint
            foreach (DigitalSignature loadedSignature in loadedSignatures)
            {
                // Get the certificate used for this signature
                X509Certificate2 loadedCert = loadedSignature.Certificate;
                string loadedThumbprint = loadedCert?.Thumbprint ?? "No certificate";

                Console.WriteLine("Loaded Signature Thumbprint: " + loadedThumbprint);

                // Verify integrity by comparing thumbprints
                bool isThumbprintMatch = string.Equals(originalThumbprint, loadedThumbprint, StringComparison.OrdinalIgnoreCase);
                Console.WriteLine("Thumbprint match: " + isThumbprintMatch);
            }
        }
    }
}
