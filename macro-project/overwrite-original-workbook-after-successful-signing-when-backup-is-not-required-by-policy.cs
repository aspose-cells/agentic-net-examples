// Title: Overwrite Original Excel Workbook with a Digital Signature using Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to load an existing XLSX file, apply a digital signature from a PFX certificate, and save the workbook back to the same path, overwriting the original file. It also shows how to verify the IsDigitallySigned flag after saving, ensuring the signature was applied successfully without creating a backup copy.
// Keywords: Aspose.Cells digital signature C# | sign Excel workbook with PFX | overwrite signed Excel file | C# Aspose.Cells save without backup | verify Excel digital signature | .NET Excel security compliance
// Common Searches: How to add a digital signature to an existing Excel file using Aspose.Cells and overwrite it | Aspose.Cells C# save signed workbook without creating a backup copy | Verify digital signature on an Excel workbook after saving with Aspose.Cells | Overwrite original Excel file after signing with a PFX certificate in .NET | C# code to sign and replace an Excel workbook using Aspose.Cells
// Developer Intent: Apply a PFX‑based digital signature to an existing Excel workbook and replace the original file only after a successful signing operation.
// Use Cases: Automate compliance‑driven signing of financial reports, overwriting the source file to meet corporate policy. | Integrate digital signing into a CI/CD pipeline that processes batches of Excel files, ensuring each file is saved in place after verification. | Secure confidential spreadsheets before distribution by signing them in‑place and confirming the signature status programmatically.
// AI Prompts: Generate C# code that uses Aspose.Cells to load an XLSX file, apply a digital signature from a PFX certificate, and overwrite the original workbook. | Suggest robust error‑handling patterns for signing an Excel workbook with Aspose.Cells, ensuring the file is only overwritten after a successful signature. | Explain how to programmatically verify the IsDigitallySigned property after saving a signed workbook with Aspose.Cells and handle verification failures.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    // This example demonstrates how to load an existing XLSX file, apply a digital signature from a PFX certificate, and save the workbook back to the same path, overwriting the original file. It also shows how to verify the IsDigitallySigned flag after saving, ensuring the signature was applied successfully without creating a backup copy.
    public class OverwriteSignedWorkbook
    {
        public static void Run()
        {
            try
            {
                // Path to the existing workbook that needs to be signed
                string workbookPath = "OriginalWorkbook.xlsx";

                // Verify workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(workbookPath);

                // Path to the certificate file (PFX)
                string certPath = "certificate.pfx";
                string certPassword = "yourPassword";

                // Verify certificate file exists
                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                // Load the certificate bytes
                byte[] certData = File.ReadAllBytes(certPath);

                // Create a digital signature using the certificate data
                DigitalSignature digitalSignature = new DigitalSignature(
                    certData,          // certificate bytes
                    certPassword,     // certificate password
                    "Signed by Aspose", // description
                    DateTime.Now);    // signing time

                // Prepare a collection and add the signature
                DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
                signatureCollection.Add(digitalSignature);

                // Add the digital signature to the workbook
                workbook.AddDigitalSignature(signatureCollection);

                // Overwrite the original workbook file with the signed version
                workbook.Save(workbookPath, SaveFormat.Xlsx);

                // Verify that the workbook is now digitally signed
                Workbook verifyWorkbook = new Workbook(workbookPath);
                Console.WriteLine("Workbook is digitally signed: " + verifyWorkbook.IsDigitallySigned);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            OverwriteSignedWorkbook.Run();
        }
    }
}
