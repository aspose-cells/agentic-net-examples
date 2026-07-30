// Title: Add a Second Digital Signature to an Existing Excel Workbook with Aspose.Cells for .NET
// Description: Demonstrates how to load an Excel file that already contains a digital signature, append a new signature using a second X509Certificate2, save the workbook, and verify that both signatures are retained and valid with Aspose.Cells.
// Keywords: Aspose.Cells add digital signature .NET | multiple digital signatures Excel | C# Aspose.Cells DigitalSignatureCollection | verify Excel digital signatures | append second signature workbook
// Common Searches: how to add another digital signature to a signed Excel file using Aspose.Cells | Aspose.Cells multiple signatures verification C# | append digital signature to existing workbook Aspose.Cells | check if both signatures persist after saving Excel with Aspose
// Developer Intent: Append a second digital signature to a workbook that already has one and confirm that both signatures are saved and valid.
// Use Cases: Load a signed workbook and retrieve its DigitalSignatureCollection. | Create a DigitalSignature from a second X509Certificate2 and add it to the collection. | Save the workbook and reload it to enumerate all signatures, displaying comments, timestamps, and validation status.
// AI Prompts: Generate C# code that uses Aspose.Cells to add a new digital signature to an already signed Excel workbook and then list all signatures. | Explain how Aspose.Cells manages multiple digital signatures in a DigitalSignatureCollection and which properties can be used for validation. | Provide best‑practice error handling for loading certificate files and verifying signatures in the described workflow.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// Demonstrates how to load an Excel file that already contains a digital signature, append a new signature using a second X509Certificate2, save the workbook, and verify that both signatures are retained and valid with Aspose.Cells.
class AddSecondSignatureDemo
{
    static void Main()
    {
        try
        {
            // Path to the workbook that already contains one digital signature
            string sourcePath = "SignedWorkbook1.xlsx";

            // Paths and passwords for the certificates
            string certPath1 = "cert1.pfx"; // first certificate (already used)
            string certPath2 = "cert2.pfx"; // second certificate to add
            string password1 = "password1";
            string password2 = "password2";

            // Output path for the workbook after adding the second signature
            string outputPath = "SignedWorkbook_WithTwoSignatures.xlsx";

            // Verify required files exist
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source workbook not found: {sourcePath}");
                return;
            }
            if (!File.Exists(certPath2))
            {
                Console.WriteLine($"Certificate file not found: {certPath2}");
                return;
            }

            // Load the already signed workbook
            Workbook workbook = new Workbook(sourcePath);

            // Retrieve the existing digital signature collection; create a new one if none exist
            DigitalSignatureCollection signatures = workbook.GetDigitalSignature();
            if (signatures == null)
            {
                signatures = new DigitalSignatureCollection();
            }

            // Load the second certificate and create a new digital signature
            X509Certificate2 cert2 = new X509Certificate2(certPath2, password2);
            DigitalSignature secondSignature = new DigitalSignature(cert2, "Second signature", DateTime.Now);

            // Add the new signature to the collection
            signatures.Add(secondSignature);

            // Apply the updated collection back to the workbook
            workbook.SetDigitalSignature(signatures);

            // Save the workbook; both signatures should now be persisted
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Reload the saved workbook to verify that both signatures are present
            Workbook verifyWorkbook = new Workbook(outputPath);
            DigitalSignatureCollection verifySignatures = verifyWorkbook.GetDigitalSignature();

            int count = 0;
            if (verifySignatures != null)
            {
                foreach (DigitalSignature sig in verifySignatures)
                {
                    count++;
                    Console.WriteLine($"Signature {count}: Comments = {sig.Comments}, SignTime = {sig.SignTime}, IsValid = {sig.IsValid}");
                }
            }

            Console.WriteLine($"Total digital signatures after adding second: {count}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
