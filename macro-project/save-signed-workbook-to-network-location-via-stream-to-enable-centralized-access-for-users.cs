// Title: Save a digitally signed Excel workbook to a UNC share via MemoryStream (Aspose.Cells for .NET)
// Description: Creates a Workbook, optionally applies a PFX‑based digital signature, ensures the target UNC directory exists, saves the workbook to a MemoryStream in XLSX format, and streams the file directly to a network share. Includes error handling and resource cleanup.
// Keywords: Aspose.Cells | C# | digital signature | PFX certificate | UNC path | network share | MemoryStream | save workbook | Excel file | .NET | FileStream
// Common Searches: Aspose.Cells save signed workbook to network share | C# write Excel to UNC path using MemoryStream | How to add a digital signature to an Excel file with Aspose.Cells | Save Excel to shared folder without temporary file | Create signed XLSX on server with Aspose.Cells
// Developer Intent: Store a digitally signed Excel file directly on a network share using Aspose.Cells and a memory stream.
// Use Cases: Generate a signed financial report and place it in a central UNC folder for team access. | Automate compliance document creation on a server and write the signed workbook to a shared drive without local temp files. | Integrate a web service that signs workbooks and saves them to a network location for downstream processing. | Batch‑process multiple workbooks, apply a PFX signature, and stream each file to a common network repository.
// AI Prompts: Write C# code that loads a .pfx certificate, signs an Aspose.Cells workbook, and saves it to a UNC path using a MemoryStream. | Show how to verify or create the target network directory and copy a MemoryStream containing an XLSX workbook to that location with proper error handling. | Provide a reusable method that accepts a Workbook, certificate path, password, and UNC destination, applies the digital signature, and returns a success flag.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// Creates a Workbook, optionally applies a PFX‑based digital signature, ensures the target UNC directory exists, saves the workbook to a MemoryStream in XLSX format, and streams the file directly to a network share. Includes error handling and resource cleanup.
class SaveSignedWorkbookToNetwork
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];
            ws.Cells["A1"].PutValue("Signed Data");

            // Load a PFX certificate for digital signing (if it exists)
            string certPath = @"C:\certs\sample.pfx";
            string certPassword = "123456";

            if (File.Exists(certPath))
            {
                byte[] certData = File.ReadAllBytes(certPath);
                DigitalSignature signature = new DigitalSignature(certData, certPassword, "Demo Signer", DateTime.Now);
                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                signatures.Add(signature);
                workbook.SetDigitalSignature(signatures);
            }
            else
            {
                Console.WriteLine($"Certificate file not found: {certPath}. Workbook will be saved without a digital signature.");
            }

            // Define the network (UNC) path where the file will be stored
            string networkPath = @"\\Server\SharedFolder\SignedWorkbook.xlsx";
            string networkDir = Path.GetDirectoryName(networkPath);
            if (!Directory.Exists(networkDir))
            {
                Directory.CreateDirectory(networkDir);
            }

            // Save the signed workbook to a memory stream in XLSX format
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0; // Reset stream position for reading

                // Write the stream content to the network location
                using (FileStream networkStream = new FileStream(networkPath, FileMode.Create, FileAccess.Write))
                {
                    ms.CopyTo(networkStream);
                }
            }

            // Release resources
            workbook.Dispose();
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
