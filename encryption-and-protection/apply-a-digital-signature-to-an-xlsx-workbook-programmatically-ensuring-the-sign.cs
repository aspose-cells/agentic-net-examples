using System;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and add sample data
            var workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].PutValue("Document requiring digital signature");

            // Generate a self‑signed certificate for demonstration purposes
            using var rsa = RSA.Create(2048);
            var req = new CertificateRequest(
                "CN=AsposeDemoCertificate",
                rsa,
                HashAlgorithmName.SHA256,
                RSASignaturePadding.Pkcs1);

            var notBefore = DateTimeOffset.UtcNow.AddDays(-1);
            var notAfter = notBefore.AddYears(1);
            var cert = req.CreateSelfSigned(notBefore, notAfter);

            // Create a digital signature using the certificate
            var signature = new DigitalSignature(cert, "Approved by QA", DateTime.UtcNow);

            // Add the signature to a collection and embed it into the workbook
            var signatureCollection = new DigitalSignatureCollection();
            signatureCollection.Add(signature);
            workbook.AddDigitalSignature(signatureCollection);

            // Save the signed workbook
            const string filePath = "SignedWorkbook.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Verify that the workbook is digitally signed
            var signedWorkbook = new Workbook(filePath);
            Console.WriteLine("Is workbook digitally signed? " + signedWorkbook.IsDigitallySigned);

            // Retrieve and display signature details
            DigitalSignatureCollection retrievedSignatures = signedWorkbook.GetDigitalSignature();
            foreach (DigitalSignature ds in retrievedSignatures)
            {
                Console.WriteLine($"Comment: {ds.Comments}");
                Console.WriteLine($"Signed at (UTC): {ds.SignTime}");
                Console.WriteLine($"Signature valid: {ds.IsValid}");
            }
        }
    }
}