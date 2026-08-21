// Title: Add a Second Digital Signature to an Already Signed Excel Workbook Using Aspose.Cells for .NET
// Description: This C# example shows how to load an Excel file that already contains a digital signature, create a new DigitalSignature from a PFX certificate, add it to the existing DigitalSignatureCollection, save the workbook, and then reload it to confirm that both signatures are retained and valid. The code demonstrates signature preservation, counting, and detailed output of each signature's comment, timestamp, and validation status.
// Keywords: Aspose.Cells | digital signature | multiple signatures | add second signature | verify Excel signatures | C# | .NET | Excel workbook signing | DigitalSignatureCollection | PFX certificate | preserve existing signatures | Xlsx digital signature
// Common Searches: how to add another digital signature to a signed Excel file using Aspose.Cells | Aspose.Cells C# add multiple digital signatures to .xlsx | verify multiple digital signatures in an Excel workbook .NET | preserve existing signature when adding a new one Aspose.Cells | count digital signatures in an Excel file with Aspose.Cells
// Developer Intent: Add a second digital signature to a workbook that is already signed and confirm that both signatures remain after saving.
// Use Cases: Create a contract workbook signed by a client, then programmatically append a manager’s signature for multi‑party approval. | Add an audit‑trail timestamp signature to a financial report after the accounting department’s signature, while keeping the original signature intact. | Automate compliance checks by enumerating and displaying all digital signatures in a workbook before distribution.
// AI Prompts: Generate C# code with Aspose.Cells to add a third digital signature to an already signed Excel workbook and return the total signature count. | Explain how to validate each signature’s IsValid property after loading a signed workbook with Aspose.Cells. | Provide a method to remove a specific digital signature from an Excel file while preserving the remaining signatures using Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    // This C# example shows how to load an Excel file that already contains a digital signature, create a new DigitalSignature from a PFX certificate, add it to the existing DigitalSignatureCollection, save the workbook, and then reload it to confirm that both signatures are retained and valid. The code demonstrates signature preservation, counting, and detailed output of each signature's comment, timestamp, and validation status.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths for the original signed workbook, the certificate, and the output workbook
                string signedWorkbookPath = "SignedWorkbook.xlsx";          // already contains one signature
                string certificatePath = "certificate.pfx";                // certificate file
                string certificatePassword = "password";                   // certificate password
                string outputWorkbookPath = "SignedWorkbook_WithTwoSignatures.xlsx";

                // Verify input files exist
                if (!File.Exists(signedWorkbookPath))
                {
                    Console.WriteLine($"Error: Workbook file not found: {signedWorkbookPath}");
                    return;
                }

                if (!File.Exists(certificatePath))
                {
                    Console.WriteLine($"Error: Certificate file not found: {certificatePath}");
                    return;
                }

                // Load the already signed workbook
                Workbook workbook = new Workbook(signedWorkbookPath);
                Console.WriteLine("Initially digitally signed: " + workbook.IsDigitallySigned);

                // Load the certificate to be used for the second signature
                X509Certificate2 certificate = new X509Certificate2(certificatePath, certificatePassword, X509KeyStorageFlags.MachineKeySet);

                // Create the second digital signature
                DigitalSignature secondSignature = new DigitalSignature(
                    certificate,
                    "Second signature added by Aspose.Cells",
                    DateTime.Now);

                // Add the new signature to the workbook (existing signatures are preserved)
                DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
                signatureCollection.Add(secondSignature);
                workbook.AddDigitalSignature(signatureCollection);

                // Save the workbook – the updated collection (now containing two signatures) is persisted
                workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);

                // Load the saved workbook to confirm both signatures are present
                Workbook verificationWorkbook = new Workbook(outputWorkbookPath);
                DigitalSignatureCollection loadedSignatures = verificationWorkbook.GetDigitalSignature();

                Console.WriteLine("After adding second signature, digitally signed: " + verificationWorkbook.IsDigitallySigned);
                Console.WriteLine("Number of digital signatures present: " + (loadedSignatures != null ? CountSignatures(loadedSignatures) : 0));

                // Display details of each signature
                if (loadedSignatures != null)
                {
                    int index = 1;
                    foreach (DigitalSignature sig in loadedSignatures)
                    {
                        Console.WriteLine($"Signature {index}:");
                        Console.WriteLine($"  Comments : {sig.Comments}");
                        Console.WriteLine($"  Sign Time: {sig.SignTime}");
                        Console.WriteLine($"  Is Valid : {sig.IsValid}");
                        index++;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Helper method to count signatures in the collection
        private static int CountSignatures(DigitalSignatureCollection collection)
        {
            int count = 0;
            foreach (DigitalSignature _ in collection)
            {
                count++;
            }
            return count;
        }
    }
}
