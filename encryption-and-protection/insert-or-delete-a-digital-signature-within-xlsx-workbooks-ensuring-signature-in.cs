using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    class Program
    {
        static void Main()
        {
            string sourcePath = "SourceWorkbook.xlsx";
            string signedPath = "SignedWorkbook.xlsx";
            string unsignedPath = "UnsignedWorkbook.xlsx";
            string certPath = "certificate.pfx";
            string certPassword = "yourPassword";

            Workbook workbook;
            if (File.Exists(sourcePath))
            {
                workbook = new Workbook(sourcePath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample data for digital signature");
                workbook.Save(sourcePath, SaveFormat.Xlsx);
            }

            X509Certificate2 certificate;
            if (File.Exists(certPath))
            {
                certificate = new X509Certificate2(certPath, certPassword, X509KeyStorageFlags.Exportable);
            }
            else
            {
                using (RSA rsa = RSA.Create(2048))
                {
                    var req = new CertificateRequest("cn=AsposeDemo", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                    certificate = req.CreateSelfSigned(DateTimeOffset.Now.AddDays(-1), DateTimeOffset.Now.AddYears(1));
                    byte[] pfx = certificate.Export(X509ContentType.Pfx, certPassword);
                    certificate = new X509Certificate2(pfx, certPassword, X509KeyStorageFlags.Exportable);
                }
            }

            DigitalSignature signature = new DigitalSignature(certificate, "Document approved", DateTime.UtcNow);
            DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
            signatureCollection.Add(signature);
            workbook.AddDigitalSignature(signatureCollection);
            workbook.Save(signedPath, SaveFormat.Xlsx);

            Workbook signedWorkbook = new Workbook(signedPath);
            Console.WriteLine("Is workbook digitally signed? " + signedWorkbook.IsDigitallySigned);

            DigitalSignatureCollection retrievedSignatures = signedWorkbook.GetDigitalSignature();
            if (retrievedSignatures != null)
            {
                foreach (DigitalSignature ds in retrievedSignatures)
                {
                    Console.WriteLine($"Comment: {ds.Comments}");
                    Console.WriteLine($"Signed at (UTC): {ds.SignTime}");
                    Console.WriteLine($"Is valid: {ds.IsValid}");
                }
            }

            signedWorkbook.RemoveDigitalSignature();
            signedWorkbook.Save(unsignedPath, SaveFormat.Xlsx);

            Workbook unsignedWorkbook = new Workbook(unsignedPath);
            Console.WriteLine("After removal, is workbook digitally signed? " + unsignedWorkbook.IsDigitallySigned);
        }
    }
}