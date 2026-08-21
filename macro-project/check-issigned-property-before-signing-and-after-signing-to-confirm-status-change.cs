// Title: Check IsDigitallySigned Before and After Signing an Excel Workbook with Aspose.Cells (C#)
// Description: Creates a new workbook, reads the IsDigitallySigned flag, loads an X509 certificate, applies a DigitalSignature via Aspose.Cells, saves to a MemoryStream, reloads the file, and verifies that the IsDigitallySigned property switches from false to true.
// Keywords: Aspose.Cells | IsDigitallySigned | C# digital signature | Excel workbook signing | X509Certificate2 | SetDigitalSignature | verify signature persistence | prevent duplicate signing
// Common Searches: how to check IsDigitallySigned in Aspose.Cells | C# verify Excel workbook digital signature after save | Aspose.Cells sign workbook with X509 certificate | detect if Excel file is already signed using Aspose
// Developer Intent: Validate that applying a digital signature changes the workbook's IsDigitallySigned status from false to true.
// Use Cases: Skip signing when a workbook is already signed to avoid duplicate signatures. | Confirm that a signature survives serialization by reloading the saved file. | Branch workflow logic based on the signed/unsigned state of an Excel document.
// AI Prompts: Write C# code that loads an X509 .pfx file, signs an Aspose.Cells workbook, and prints IsDigitallySigned before and after saving. | Show how to reload a signed workbook from a MemoryStream and confirm the digital signature using IsDigitallySigned. | Provide robust error handling for missing certificate files and invalid passwords when using Aspose.Cells digital signatures.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsSignatureCheck
{
    // Creates a new workbook, reads the IsDigitallySigned flag, loads an X509 certificate, applies a DigitalSignature via Aspose.Cells, saves to a MemoryStream, reloads the file, and verifies that the IsDigitallySigned property switches from false to true.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Digital Signature Test");

                // Check IsDigitallySigned before signing
                bool isSignedBefore = workbook.IsDigitallySigned;
                Console.WriteLine("Is workbook digitally signed before signing? " + isSignedBefore);

                // Load certificate (replace with a valid .pfx file and password)
                string certPath = "test.pfx";
                string certPassword = "password";

                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);

                // Create a digital signature
                DigitalSignature signature = new DigitalSignature(certificate, "Demo Signature", DateTime.Now);

                // Add signature to a collection and apply to workbook
                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                signatures.Add(signature);
                workbook.SetDigitalSignature(signatures);

                // Save the signed workbook to a memory stream
                using (MemoryStream signedStream = new MemoryStream())
                {
                    workbook.Save(signedStream, SaveFormat.Xlsx);
                    signedStream.Position = 0;

                    // Reload workbook from the stream to verify signature persistence
                    Workbook signedWorkbook = new Workbook(signedStream);
                    bool isSignedAfter = signedWorkbook.IsDigitallySigned;
                    Console.WriteLine("Is workbook digitally signed after signing? " + isSignedAfter);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
