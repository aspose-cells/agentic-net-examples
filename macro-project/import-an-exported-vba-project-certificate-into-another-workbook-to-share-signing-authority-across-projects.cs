// Title: Import and Apply a VBA Project Certificate to Another Workbook with Aspose.Cells (.NET)
// Description: Load a signed workbook, extract its VBA project's CertRawData, create a DigitalSignature, and sign a different macro‑enabled workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells VBA certificate import | C# copy VBA digital signature | sign macro-enabled workbook programmatically | VbaProject CertRawData extraction | Excel VBA digital signature .NET | import VBA certificate Aspose | automated Excel macro signing
// Common Searches: how to copy a VBA certificate from one Excel file to another using Aspose.Cells | C# code to extract CertRawData from a signed VBA project | programmatically sign a .xlsm workbook with an existing VBA certificate | Aspose.Cells import VBA project certificate example | digital signature for Excel macro projects in .NET
// Developer Intent: Extract a VBA project's certificate from a signed workbook and reuse it to sign another workbook's VBA project.
// Use Cases: Standardize corporate VBA signing across multiple macro‑enabled reports. | Automate signing of generated Excel files that contain macros, preserving the original authority. | Migrate existing signed VBA projects to new workbooks without losing the digital signature.
// AI Prompts: Write C# code that reads CertRawData from a signed VBA project with Aspose.Cells and signs another workbook's VBA project using the same certificate. | Explain how to handle password‑protected VBA certificates when importing them via Aspose.Cells, including error handling best practices. | Provide a step‑by‑step guide to verify a source VBA project is signed before extracting its certificate and applying it to a target workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaCertificateImport
{
    // Load a signed workbook, extract its VBA project's CertRawData, create a DigitalSignature, and sign a different macro‑enabled workbook using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the workbook that already contains a signed VBA project (source)
                string sourcePath = "SourceSigned.xlsm";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: '{sourcePath}'.");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Access its VBA project
                VbaProject sourceVba = sourceWorkbook.VbaProject;

                // Ensure the source VBA project is signed
                if (!sourceVba.IsSigned)
                {
                    Console.WriteLine("Source VBA project is not signed. Exiting.");
                    return;
                }

                // Retrieve the raw certificate data from the signed VBA project
                byte[] certificateRawData = sourceVba.CertRawData;

                // Password used to protect the original certificate (replace with actual password)
                string certificatePassword = "your_certificate_password";

                // Path to the target workbook that will receive the same signing authority
                string targetPath = "TargetUnsigned.xlsm";

                // Verify target file exists
                if (!File.Exists(targetPath))
                {
                    Console.WriteLine($"Target file not found: '{targetPath}'.");
                    return;
                }

                // Load the target workbook (it may or may not already have a VBA project)
                Workbook targetWorkbook = new Workbook(targetPath);

                // Access the VBA project of the target workbook
                VbaProject targetVba = targetWorkbook.VbaProject;

                // Create a DigitalSignature object using the raw certificate data
                DigitalSignature importedSignature = new DigitalSignature(
                    certificateRawData,          // certificate bytes
                    certificatePassword,        // password for the certificate
                    "Imported VBA Signature",   // optional comment
                    DateTime.Now);              // signing time

                // Sign the target VBA project with the imported certificate
                targetVba.Sign(importedSignature);

                // Save the signed target workbook
                string outputPath = "TargetSigned.xlsm";
                targetWorkbook.Save(outputPath, SaveFormat.Xlsm);

                Console.WriteLine($"Target workbook signed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
