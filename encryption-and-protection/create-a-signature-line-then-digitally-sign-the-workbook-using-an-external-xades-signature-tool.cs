// Title: Create a Signature Line and Apply an XAdES Digital Signature to an Excel Workbook with Aspose.Cells (.NET)
// Description: Demonstrates how to add a SignatureLine shape to a worksheet, load a PFX certificate, generate an XAdES digital signature, link the signature line and signature via a shared Id, add the signature to a DigitalSignatureCollection, and save the workbook while gracefully handling missing certificates.
// Keywords: Aspose.Cells signature line | XAdES digital signature .NET | sign Excel workbook C# | PFX certificate Aspose.Cells | DigitalSignatureCollection | link signature line to digital signature | Excel workbook protection
// Common Searches: how to add a signature line in Excel using Aspose.Cells | apply XAdES signature to a workbook with Aspose.Cells .NET | link signature line ID to digital signature Aspose.Cells | save Excel file with digital signature only if certificate exists | Aspose.Cells C# digital signature example
// Developer Intent: Insert a signature line into a worksheet and digitally sign the workbook with an XAdES signature, ensuring the line and signature are linked.
// Use Cases: Add a signature line at a specific cell for an approval workflow and sign the file with a PFX certificate. | Generate a signed Excel report only when the required certificate file is present; otherwise, save an unsigned version. | Apply multiple XAdES signatures to a workbook, each tied to its own signature line via unique identifiers.
// AI Prompts: Write C# code that places a signature line at row 10, column 3 and signs the workbook with an XAdES signature using a given PFX file. | Explain how to detect a missing certificate file and continue saving the workbook without a digital signature in Aspose.Cells. | Show how to add several XAdES signatures to a workbook, linking each signature line to its corresponding digital signature.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsSignatureDemo
{
    // Demonstrates how to add a SignatureLine shape to a worksheet, load a PFX certificate, generate an XAdES digital signature, link the signature line and signature via a shared Id, add the signature to a DigitalSignatureCollection, and save the workbook while gracefully handling missing certificates.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure a signature line
                SignatureLine sigLine = new SignatureLine
                {
                    Signer = "John Doe",
                    Title = "Approver",
                    Email = "john.doe@example.com",
                    Instructions = "Please sign to approve the document.",
                    AllowComments = true,
                    ShowSignedDate = true,
                    IsLine = true
                };

                // Add the signature line to the worksheet at row 5, column 2 (zero‑based indexes)
                Picture picture = worksheet.Shapes.AddSignatureLine(5, 2, sigLine);

                // Path to the signing certificate (PFX file)
                string certPath = "myCertificate.pfx";
                string certPassword = "password123";

                DigitalSignatureCollection signatureCollection = null;

                // Load certificate and create digital signature only if the file exists
                if (File.Exists(certPath))
                {
                    try
                    {
                        byte[] certData = File.ReadAllBytes(certPath);
                        DigitalSignature digitalSignature = new DigitalSignature(certData, certPassword, "Document approved", DateTime.Now);
                        digitalSignature.XAdESType = XAdESType.XAdES;

                        // Link the signature line and digital signature by Id
                        sigLine.Id = Guid.NewGuid();
                        picture.SignatureLine.Id = sigLine.Id;
                        digitalSignature.Id = sigLine.Id;

                        // Prepare collection and apply to workbook
                        signatureCollection = new DigitalSignatureCollection();
                        signatureCollection.Add(digitalSignature);
                        workbook.SetDigitalSignature(signatureCollection);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error creating digital signature: {ex.Message}");
                        Console.WriteLine("The workbook will be saved without a digital signature.");
                    }
                }
                else
                {
                    Console.WriteLine("Certificate file not found. The workbook will be saved without a digital signature.");
                }

                // Save the workbook
                string outputPath = "SignedWorkbook_WithSignatureLine.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
