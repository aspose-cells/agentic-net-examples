// Title: Create a Signature Line and Apply an XAdES Digital Signature to an Excel Workbook using Aspose.Cells for .NET
// Description: Shows how to add a configurable signature line at cell B2, load a PFX certificate, generate an XAdES‑type DigitalSignature, attach it through a DigitalSignatureCollection, and save the signed file (SignedWorkbook.xlsx) with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# signature line | Excel digital signature | XAdES | PFX certificate | DigitalSignatureCollection | external XAdES tool | sign workbook | add signature line | Aspose.Cells .NET
// Common Searches: Aspose.Cells add signature line C# | XAdES digital signature Excel .NET | How to sign an Excel workbook with a PFX certificate | Create signature line in Excel using Aspose.Cells | Apply external XAdES signature to workbook
// Developer Intent: Insert a signature line into a worksheet and sign the workbook with an XAdES signature using a PFX certificate.
// Use Cases: Place a signature line at a specific cell (e.g., B2) and define signer details such as name, title, and email. | Load a .pfx certificate, create an XAdES DigitalSignature, add it to a DigitalSignatureCollection, and embed the collection into the workbook. | Generate a signed Excel file that complies with XAdES standards for legal and compliance scenarios.
// AI Prompts: Write C# code that adds a signature line to cell C3 and signs the workbook with an XAdES signature stored in Azure Key Vault. | Explain how to validate an XAdES digital signature in an Excel file created with Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.DigitalSignatures;

// Shows how to add a configurable signature line at cell B2, load a PFX certificate, generate an XAdES‑type DigitalSignature, attach it through a DigitalSignatureCollection, and save the signed file (SignedWorkbook.xlsx) with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure a signature line
        SignatureLine signatureLine = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Manager",
            Email = "john.doe@example.com",
            Instructions = "Please sign here.",
            AllowComments = true,
            ShowSignedDate = true,
            IsLine = true
        };

        // Add the signature line to the worksheet at cell B2 (row index 1, column index 1)
        worksheet.Shapes.AddSignatureLine(1, 1, signatureLine);

        // Load the certificate (replace with actual path and password)
        string certPath = "mycert.pfx";
        string certPassword = "password";

        // Create a digital signature using the certificate byte array (Bouncy Castle constructor)
        byte[] certData = File.ReadAllBytes(certPath);
        DigitalSignature digitalSignature = new DigitalSignature(certData, certPassword, "Signed by external XAdES tool", DateTime.UtcNow);

        // Indicate that this is an XAdES signature
        digitalSignature.XAdESType = XAdESType.XAdES;

        // Add the signature to a collection
        DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
        signatureCollection.Add(digitalSignature);

        // Apply the digital signature to the workbook
        workbook.SetDigitalSignature(signatureCollection);

        // Save the signed workbook
        workbook.Save("SignedWorkbook.xlsx");
    }
}
