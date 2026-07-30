// Title: Sign and Verify a Protected Worksheet with Aspose.Cells for .NET
// Description: Creates a workbook, protects the first worksheet, loads an X509 PFX certificate, builds a DigitalSignature with comment and UTC timestamp, adds it to a DigitalSignatureCollection, applies the signature, saves the file, reloads it, checks IsDigitallySigned, retrieves the signature collection, and displays each signature's comment, sign time and validation status.
// Keywords: Aspose.Cells | digital signature | protected worksheet | C# | .NET | X509 certificate | PFX file | Workbook signing | signature verification | IsDigitallySigned | compliance | tamper evidence
// Common Searches: Aspose.Cells sign protected worksheet C# | Verify Excel workbook digital signature with Aspose.Cells | How to use X509 certificate to sign an Aspose.Cells workbook | Check IsDigitallySigned property after protecting a sheet | DigitalSignatureCollection example in Aspose.Cells .NET
// Developer Intent: Add a digital signature to a worksheet that is already protected and confirm that the signature is valid after the workbook is saved.
// Use Cases: Secure confidential Excel reports by protecting sheets and adding a tamper‑evident digital signature. | Automate compliance checks that validate incoming signed workbooks before processing. | Integrate signing and verification into a document‑generation pipeline that logs signature details for audit trails.
// AI Prompts: Generate C# code that protects a worksheet, signs it with a PFX certificate using Aspose.Cells, and validates the signature after saving. | Explain how to handle a DigitalSignature.IsValid failure when loading a signed workbook with Aspose.Cells. | Provide a step‑by‑step guide to protect a worksheet, apply a digital signature, save the workbook, and retrieve signature information in Aspose.Cells for .NET.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// Creates a workbook, protects the first worksheet, loads an X509 PFX certificate, builds a DigitalSignature with comment and UTC timestamp, adds it to a DigitalSignatureCollection, applies the signature, saves the file, reloads it, checks IsDigitallySigned, retrieves the signature collection, and displays each signature's comment, sign time and validation status.
class DigitalSignatureProtectedWorksheetDemo
{
    static void Main()
    {
        // Path to the certificate file and its password
        string certPath = "mycert.pfx";
        string certPassword = "password";

        // ---------- Create a new workbook ----------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "ProtectedSheet";

        // Add some data to the worksheet
        sheet.Cells["A1"].PutValue("Confidential Data");

        // Protect the worksheet (all protection types)
        sheet.Protect(ProtectionType.All);

        // ---------- Load the signing certificate ----------
        X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);

        // ---------- Create a digital signature ----------
        DigitalSignature signature = new DigitalSignature(
            certificate,               // certificate used for signing
            "Worksheet signed",        // comment / purpose
            DateTime.UtcNow);          // signing time (UTC)

        // ---------- Add the signature to a collection ----------
        DigitalSignatureCollection signatures = new DigitalSignatureCollection();
        signatures.Add(signature);

        // ---------- Apply the digital signature to the workbook ----------
        workbook.SetDigitalSignature(signatures);

        // ---------- Save the signed workbook ----------
        string signedPath = "SignedProtectedWorkbook.xlsx";
        workbook.Save(signedPath, SaveFormat.Xlsx);

        // ---------- Load the signed workbook for verification ----------
        Workbook signedWorkbook = new Workbook(signedPath);

        // Check if the workbook reports being digitally signed
        Console.WriteLine("Workbook IsDigitallySigned: " + signedWorkbook.IsDigitallySigned);

        // Retrieve the digital signature collection
        DigitalSignatureCollection loadedSignatures = signedWorkbook.GetDigitalSignature();

        // Verify each signature's validity and display its details
        foreach (DigitalSignature ds in loadedSignatures)
        {
            Console.WriteLine($"Comment   : {ds.Comments}");
            Console.WriteLine($"Sign Time : {ds.SignTime.ToUniversalTime()} (UTC)");
            Console.WriteLine($"Is Valid  : {ds.IsValid}");
        }
    }
}
