// Title: Import a VBA Project Certificate and Sign Another Workbook with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to load a signed macro‑enabled workbook, extract its VBA project certificate via VbaProject.CertRawData, create an X509Certificate2 and DigitalSignature, and apply the same certificate to sign a different workbook's VBA project using Aspose.Cells.
// Keywords: Aspose.Cells VBA certificate import | C# copy VBA digital signature | VbaProject.CertRawData | DigitalSignature Aspose.Cells | sign macro-enabled workbook programmatically | X509Certificate2 Excel VBA | batch sign Excel macros | Excel VBA project signing
// Common Searches: how to copy a VBA project certificate between Excel files with Aspose.Cells | C# extract VBA certificate raw data from signed workbook | use Aspose.Cells to sign another workbook with an existing VBA certificate | import VBA digital signature in .NET | Aspose.Cells VbaProject.Sign example
// Developer Intent: Retrieve a VBA project's digital certificate from a signed workbook and reuse it to sign a different workbook's VBA project programmatically.
// Use Cases: Apply a corporate VBA signing certificate across multiple macro‑enabled workbooks without user interaction. | Automate batch signing of newly generated workbooks using a trusted certificate stored in an existing file. | Consolidate several signed workbooks into a master file while preserving the original signing authority.
// AI Prompts: Generate C# code that extracts a VBA project's certificate from a signed workbook using Aspose.Cells and signs another workbook with it. | Explain how to handle password‑protected VBA certificates when importing them with Aspose.Cells and X509Certificate2. | Provide robust error‑handling patterns for signing a VBA project with an imported digital certificate in Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

// Demonstrates how to load a signed macro‑enabled workbook, extract its VBA project certificate via VbaProject.CertRawData, create an X509Certificate2 and DigitalSignature, and apply the same certificate to sign a different workbook's VBA project using Aspose.Cells.
class ImportVbaCertificate
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SignedSource.xlsm";
            const string targetPath = "Target.xlsm";
            const string outputPath = "TargetSigned.xlsm";

            // Verify source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Verify target workbook exists
            if (!File.Exists(targetPath))
            {
                Console.WriteLine($"Target file not found: {targetPath}");
                return;
            }

            // Load the workbook that already has a signed VBA project (source)
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Retrieve the raw certificate data from the source VBA project
            byte[] certificateData = sourceWorkbook.VbaProject.CertRawData;

            // Ensure the source workbook is actually signed
            if (certificateData == null || certificateData.Length == 0)
            {
                Console.WriteLine("Source workbook does not contain a signed VBA project.");
                return;
            }

            // Load the certificate from the raw data.
            // Replace "sourceCertPassword" with the actual password of the source certificate if required.
            string sourceCertPassword = "sourceCertPassword";
            X509Certificate2 certificate = new X509Certificate2(certificateData, sourceCertPassword);

            // Create a DigitalSignature object using the imported certificate
            DigitalSignature importedSignature = new DigitalSignature(
                certificate,
                "Imported VBA Signature",
                DateTime.Now);

            // Load the target workbook that will receive the imported certificate
            Workbook targetWorkbook = new Workbook(targetPath);

            // Sign the target workbook's VBA project with the imported digital signature
            VbaProject targetVbaProject = targetWorkbook.VbaProject;
            targetVbaProject.Sign(importedSignature);

            // Save the target workbook; it will now be signed with the same certificate as the source
            targetWorkbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine("Certificate imported and VBA project signed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
