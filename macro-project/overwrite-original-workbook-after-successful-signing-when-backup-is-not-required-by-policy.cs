// Title: Overwrite Existing Excel Workbook After Digital Signing with Aspose.Cells for .NET (C#)
// Description: Loads or creates an .xlsx file, reads a PFX certificate, builds a DigitalSignature, applies it via Workbook.SetDigitalSignature, and saves the workbook back to the same path using SaveFormat.Xlsx. The operation overwrites the original file without generating a backup, complying with policies that forbid duplicate copies.
// Keywords: Aspose.Cells | C# | digital signature | Excel workbook | overwrite file | no backup | SetDigitalSignature | PFX certificate | SaveFormat.Xlsx | policy‑driven saving
// Common Searches: Aspose.Cells overwrite signed workbook C# | save Excel file without backup after digital signature | apply PFX certificate to .xlsx using Aspose.Cells | C# code to sign and replace existing Excel file | how to use SetDigitalSignature in Aspose.Cells
// Developer Intent: Apply a digital signature to an Excel workbook and replace the original file in place, avoiding any backup creation.
// Use Cases: Automated compliance signing of financial reports that must remain a single file. | CI/CD pipelines that generate Excel artifacts, sign them, and overwrite the source to reduce storage. | Enterprise workflows that update spreadsheets with a corporate certificate before archiving, eliminating duplicate versions.
// AI Prompts: Write C# code using Aspose.Cells to load an existing .xlsx, sign it with a PFX certificate, and overwrite the original file without creating a backup. | Explain how to detect a missing certificate file and still ensure the workbook is saved safely. | Show how to customize the signer name and timestamp in the DigitalSignature object before overwriting the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// Loads or creates an .xlsx file, reads a PFX certificate, builds a DigitalSignature, applies it via Workbook.SetDigitalSignature, and saves the workbook back to the same path using SaveFormat.Xlsx. The operation overwrites the original file without generating a backup, complying with policies that forbid duplicate copies.
class SignWorkbookOverwrite
{
    static void Main()
    {
        try
        {
            // Path to the workbook that will be signed and overwritten
            string workbookPath = "SignedWorkbook.xlsx";

            // Load existing workbook or create a new one if it does not exist
            Workbook workbook;
            if (File.Exists(workbookPath))
            {
                workbook = new Workbook(workbookPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Save(workbookPath, SaveFormat.Xlsx);
            }

            // Path to the certificate (PFX) file and its password
            string certPath = Path.Combine("certs", "sample.pfx");
            string certPassword = "1234567890";

            // Apply digital signature only if the certificate file exists
            if (File.Exists(certPath))
            {
                // Read the certificate data into a byte array
                byte[] certData = File.ReadAllBytes(certPath);

                // Create a digital signature using the certificate data
                DigitalSignature signature = new DigitalSignature(
                    certData,          // certificate bytes
                    certPassword,     // password for the certificate
                    "Sample Signer",  // signer name
                    DateTime.Now);    // signing time

                // Add the signature to a collection (required by SetDigitalSignature)
                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                signatures.Add(signature);

                // Apply the digital signature to the workbook
                workbook.SetDigitalSignature(signatures);
            }
            else
            {
                Console.WriteLine($"Certificate file not found: {certPath}. Skipping signing.");
            }

            // Overwrite the original workbook file (no backup is created)
            workbook.Save(workbookPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {workbookPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
