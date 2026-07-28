// Title: Add a Signature Line and Digital Signature to Excel using a Thumbprint‑Loaded Certificate (Aspose.Cells C#)
// Description: This C# example shows how to retrieve an X509Certificate2 from the current user's Personal store by thumbprint, create a workbook with Aspose.Cells, insert a visible signature line, generate a matching DigitalSignature, attach it to the workbook, and save the signed XLSX file.
// Keywords: Aspose.Cells signature line | C# digital signature Excel | load certificate by thumbprint | Windows certificate store Aspose | Excel workbook signing .NET | X509Certificate2 thumbprint | AddSignatureLineWithThumbprint
// Common Searches: Aspose.Cells add signature line C# | sign Excel file with certificate thumbprint | load X509Certificate2 from Windows store in .NET | digital signature for Excel workbook Aspose | visible signature line and digital signature together
// Developer Intent: Create an Excel file, place a visible signature line, and apply a digital signature using a certificate identified by its thumbprint.
// Use Cases: Automated contract approval where the visible signature line and the underlying digital signature use a corporate certificate stored in Windows. | Generating audited financial reports that require both a signature line for reviewers and a cryptographic signature for compliance. | Batch‑signing exported spreadsheets in a Windows environment using a specific user or service certificate.
// AI Prompts: Extend the sample to enumerate all certificates in the store and let the user pick one based on thumbprint or subject name. | Add comprehensive error handling for missing, expired, or non‑exportable certificates when signing an Excel workbook with Aspose.Cells. | Show how to verify the digital signature after saving the workbook using Aspose.Cells APIs.

using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.DigitalSignatures;

// This C# example shows how to retrieve an X509Certificate2 from the current user's Personal store by thumbprint, create a workbook with Aspose.Cells, insert a visible signature line, generate a matching DigitalSignature, attach it to the workbook, and save the signed XLSX file.
class AddSignatureLineWithThumbprint
{
    static void Main()
    {
        // ----- 1. Load the certificate from Windows certificate store by thumbprint -----
        string thumbprint = "YOUR_CERTIFICATE_THUMBPRINT"; // replace with actual thumbprint (no spaces)
        X509Certificate2 certificate = null;

        X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly);
        var certs = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
        if (certs.Count > 0)
        {
            certificate = certs[0];
        }
        store.Close();

        if (certificate == null)
        {
            Console.WriteLine("Certificate with the specified thumbprint was not found.");
            return;
        }

        // ----- 2. Create a new workbook and add a signature line -----
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Configure the signature line
        SignatureLine signatureLine = new SignatureLine
        {
            Signer = "John Doe",
            Title = "Approver",
            Email = "john.doe@example.com",
            Id = Guid.NewGuid(),               // unique identifier for the line
            ProviderId = Guid.Empty,           // default provider
            SignatureLineType = SignatureType.Default,
            AllowComments = true,
            ShowSignedDate = true,
            Instructions = "Please sign here."
        };

        // Add the signature line to the worksheet (row 5, column 2 as an example)
        Picture pic = ws.Shapes.AddSignatureLine(5, 2, signatureLine);

        // ----- 3. Create a digital signature that references the same certificate -----
        DigitalSignature digitalSignature = new DigitalSignature(
            certificate,
            "Signed using certificate from thumbprint",
            DateTime.Now);

        // Associate the digital signature with the signature line via the Id
        digitalSignature.Id = signatureLine.Id;
        digitalSignature.ProviderId = signatureLine.ProviderId;

        // ----- 4. Add the digital signature to the workbook -----
        DigitalSignatureCollection dsCollection = new DigitalSignatureCollection();
        dsCollection.Add(digitalSignature);
        wb.SetDigitalSignature(dsCollection);

        // ----- 5. Save the signed workbook -----
        string outputPath = "SignedWorkbook.xlsx";
        wb.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved and signed at: {outputPath}");
    }
}
