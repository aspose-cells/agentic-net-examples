// Title: Sign and Verify a Protected Worksheet with Aspose.Cells for .NET
// Description: Demonstrates how to protect a worksheet, attach a digital signature using an X509 PFX certificate, save the workbook, and then reload it to check IsDigitallySigned, retrieve the DigitalSignatureCollection, and read each signature’s comments, timestamp and validation status.
// Keywords: Aspose.Cells digital signature | C# protect worksheet | Excel digital signature verification | X509Certificate2 signing | SetDigitalSignature | IsDigitallySigned | DigitalSignatureCollection | Workbook signing .NET | Excel compliance
// Common Searches: Aspose.Cells add digital signature to protected sheet C# | Verify Excel workbook signature after protection Aspose | How to sign an Excel file with a PFX certificate using Aspose.Cells | Check digital signature validity in a .NET workbook | Protect worksheet and apply digital signature Aspose.Cells example
// Developer Intent: Apply a digital signature to a worksheet that has been protected and programmatically confirm the signature’s authenticity.
// Use Cases: Secure financial or legal reports by protecting the sheet and signing the workbook before distribution. | Automate compliance audits by validating that received Excel files are signed and unchanged. | Integrate digital signing into a document‑generation pipeline to guarantee integrity of protected worksheets.
// AI Prompts: Generate C# code that protects an Aspose.Cells worksheet, signs it with a PFX certificate, and saves the workbook. | Show how to load a signed Excel file with Aspose.Cells and programmatically verify each digital signature’s comments, timestamp, and validity. | Explain how to handle signature verification failures and retrieve detailed error information using Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    // Demonstrates how to protect a worksheet, attach a digital signature using an X509 PFX certificate, save the workbook, and then reload it to check IsDigitallySigned, retrieve the DigitalSignatureCollection, and read each signature’s comments, timestamp and validation status.
    class Program
    {
        static void Main()
        {
            // Path to the certificate file (PFX) and its password
            string certPath = "myCertificate.pfx";
            string certPassword = "certPassword";

            // Output file for the signed workbook
            string signedFile = "ProtectedSignedWorkbook.xlsx";

            // -------------------------------------------------
            // 1. Create a new workbook and add sample data
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProtectedSheet";
            sheet.Cells["A1"].PutValue("Data to be protected and signed");

            // -------------------------------------------------
            // 2. Protect the worksheet (all protection types)
            // -------------------------------------------------
            sheet.Protect(ProtectionType.All);

            // -------------------------------------------------
            // 3. Load the certificate and create a digital signature
            // -------------------------------------------------
            X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);
            DigitalSignature signature = new DigitalSignature(certificate, "Workbook Signature", DateTime.UtcNow);

            // -------------------------------------------------
            // 4. Add the signature to a collection and set it on the workbook
            // -------------------------------------------------
            DigitalSignatureCollection signatures = new DigitalSignatureCollection();
            signatures.Add(signature);
            workbook.SetDigitalSignature(signatures);

            // -------------------------------------------------
            // 5. Save the signed workbook
            // -------------------------------------------------
            workbook.Save(signedFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved with digital signature: {signedFile}");

            // -------------------------------------------------
            // 6. Load the saved workbook and verify the signature
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook(signedFile);

            // Check if the workbook reports being digitally signed
            Console.WriteLine($"Is workbook digitally signed? {loadedWorkbook.IsDigitallySigned}");

            // Retrieve the digital signature collection
            DigitalSignatureCollection loadedSignatures = loadedWorkbook.GetDigitalSignature();

            if (loadedSignatures != null)
            {
                foreach (DigitalSignature ds in loadedSignatures)
                {
                    Console.WriteLine($"Signature Comments : {ds.Comments}");
                    Console.WriteLine($"Signature Time     : {ds.SignTime}");
                    Console.WriteLine($"Signature IsValid  : {ds.IsValid}");
                }
            }
            else
            {
                Console.WriteLine("No digital signatures found in the workbook.");
            }
        }
    }
}
