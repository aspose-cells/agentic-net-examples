// Title: Save a digitally signed Excel workbook to a UNC network share with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, loads a PFX certificate, applies a digital signature, ensures the target folder exists, and writes the signed file to a UNC path via FileStream using Aspose.Cells SaveFormat.Xlsx.
// Keywords: Aspose.Cells | C# | digital signature | PFX certificate | UNC path | network share | FileStream | SaveFormat.Xlsx | centralized Excel storage | remote folder
// Common Searches: Aspose.Cells save signed workbook to network share | C# write Excel file to UNC path with digital signature | How to apply PFX certificate to Excel using Aspose.Cells | Save Excel workbook via FileStream to remote folder | Store signed Excel reports on a shared server
// Developer Intent: Write a signed Excel file directly to a network location using a stream.
// Use Cases: Automatically generate a digitally signed financial statement and place it on a shared server for team access. | Create compliance documents each night, sign them with a corporate certificate, and store them in a central file share. | Deploy macro‑enabled workbooks that require a trusted signature to a network folder for secure distribution across the organization.
// AI Prompts: Generate C# code that loads a PFX certificate, signs an Aspose.Cells workbook, and saves it to a UNC path using FileStream. | Explain how to verify the certificate file exists and create the destination network directory before saving a signed workbook with Aspose.Cells. | Show how to adapt the example to export the signed workbook as PDF while preserving the digital signature.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// Creates a workbook, adds sample data, loads a PFX certificate, applies a digital signature, ensures the target folder exists, and writes the signed file to a UNC path via FileStream using Aspose.Cells SaveFormat.Xlsx.
class SaveSignedWorkbookToNetwork
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            using (Workbook workbook = new Workbook())
            {
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Signed Data");

                // Load a PFX certificate file (replace with your actual path and password)
                string certFilePath = @"C:\Certificates\sample.pfx";
                string certPassword = "password";

                if (File.Exists(certFilePath))
                {
                    byte[] certData = File.ReadAllBytes(certFilePath);

                    // Create a digital signature collection and add a signature
                    DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                    DigitalSignature signature = new DigitalSignature(certData, certPassword, "Demo Signer", DateTime.Now);
                    signatures.Add(signature);

                    // Apply the digital signature to the workbook
                    workbook.SetDigitalSignature(signatures);
                }
                else
                {
                    Console.WriteLine($"Certificate file not found at '{certFilePath}'. Skipping digital signature.");
                }

                // Define the UNC network path where the workbook will be saved
                string networkFilePath = @"\\Server\Share\Documents\SignedWorkbook.xlsx";

                // Ensure the target directory exists
                string networkDir = Path.GetDirectoryName(networkFilePath);
                if (!string.IsNullOrEmpty(networkDir) && !Directory.Exists(networkDir))
                {
                    Directory.CreateDirectory(networkDir);
                }

                // Save the signed workbook to the network location using a stream
                using (FileStream networkStream = new FileStream(networkFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Save(networkStream, SaveFormat.Xlsx);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
