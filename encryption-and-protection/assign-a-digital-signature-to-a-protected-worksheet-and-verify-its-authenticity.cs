// Title: C# – Sign and Verify a Password‑Protected Excel Worksheet with Aspose.Cells Digital Signatures
// Description: Shows how to protect the first worksheet of a new workbook, load an X509 .pfx certificate, create a digital signature with a comment and UTC timestamp, embed it in the workbook, save the file, then reload it to check the IsDigitallySigned flag, enumerate the signature collection, and display each signature’s comment, signing time and validation result.
// Keywords: Aspose.Cells | C# digital signature | Excel worksheet protection | X509 certificate signing | verify workbook signature | protected sheet signing | digital signature collection | IsDigitallySigned | Aspose.Cells .NET | certificate .pfx
// Common Searches: Aspose.Cells sign protected worksheet C# | Add digital signature to Excel file using Aspose | Verify Excel digital signature with Aspose.Cells | Load X509 certificate for Aspose.Cells signing | Check if workbook is digitally signed Aspose
// Developer Intent: Add a digital signature to a password‑protected Excel worksheet and confirm its authenticity programmatically.
// Use Cases: Secure distribution of confidential reports by protecting the sheet and embedding a certificate‑based signature. | Automated compliance checks that load a signed workbook, verify its integrity, and extract signature metadata. | Integrate digital signing into CI/CD pipelines to guarantee the authenticity of generated Excel deliverables.
// AI Prompts: Generate C# code with Aspose.Cells that protects the first worksheet, signs it using a .pfx certificate, saves the file, and then verifies the signature. | Explain how to retrieve and display digital‑signature comments, signing timestamps, and validation results from a signed workbook. | Provide best‑practice error‑handling for missing certificate files, incorrect passwords, and failed signature validation when using Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    // Shows how to protect the first worksheet of a new workbook, load an X509 .pfx certificate, create a digital signature with a comment and UTC timestamp, embed it in the workbook, save the file, then reload it to check the IsDigitallySigned flag, enumerate the signature collection, and display each signature’s comment, signing time and validation result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the certificate (PFX) file and its password
                string certPath = "myCertificate.pfx";
                string certPassword = "certPassword";

                // Verify that the certificate file exists
                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                // Output workbook path
                string signedWorkbookPath = "ProtectedSignedWorkbook.xlsx";

                // -------------------------------------------------
                // 1. Create a new workbook and protect its first sheet
                // -------------------------------------------------
                Workbook workbook = new Workbook();                     // create
                Worksheet sheet = workbook.Worksheets[0];

                // Protect the worksheet with a password (all protection types)
                // The third parameter is the old password; an empty string is acceptable for a new protection
                sheet.Protect(ProtectionType.All, "sheetPassword", string.Empty);

                // -------------------------------------------------
                // 2. Load the certificate and create a digital signature
                // -------------------------------------------------
                X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);
                DigitalSignature signature = new DigitalSignature(
                    certificate,                     // certificate containing private key
                    "Signed protected worksheet",    // comment
                    DateTime.UtcNow);                // sign time (UTC)

                // -------------------------------------------------
                // 3. Add the signature to a collection and apply it to the workbook
                // -------------------------------------------------
                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                signatures.Add(signature);
                workbook.SetDigitalSignature(signatures);               // set signature

                // -------------------------------------------------
                // 4. Save the signed workbook
                // -------------------------------------------------
                workbook.Save(signedWorkbookPath, SaveFormat.Xlsx);     // save
                Console.WriteLine($"Workbook saved to: {signedWorkbookPath}");

                // -------------------------------------------------
                // 5. Load the saved workbook and verify the signature
                // -------------------------------------------------
                if (!File.Exists(signedWorkbookPath))
                {
                    Console.WriteLine($"Saved workbook not found: {signedWorkbookPath}");
                    return;
                }

                Workbook loadedWorkbook = new Workbook(signedWorkbookPath); // load
                Console.WriteLine("Workbook is digitally signed: " + loadedWorkbook.IsDigitallySigned);

                DigitalSignatureCollection loadedSignatures = loadedWorkbook.GetDigitalSignature();
                if (loadedSignatures != null)
                {
                    foreach (DigitalSignature ds in loadedSignatures)
                    {
                        Console.WriteLine("Signature comment : " + ds.Comments);
                        Console.WriteLine("Signature time    : " + ds.SignTime);
                        Console.WriteLine("Signature valid   : " + ds.IsValid);
                    }
                }
                else
                {
                    Console.WriteLine("No digital signatures found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
