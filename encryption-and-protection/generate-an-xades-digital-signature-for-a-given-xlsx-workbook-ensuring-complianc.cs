using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace XAdESSignatureDemo
{
    class Program
    {
        static void Main()
        {
            // Paths and parameters
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string sourceDir = Path.Combine(baseDir, "Data", "Source");
            string outputDir = Path.Combine(baseDir, "Data", "Output");
            string sourceFile = Path.Combine(sourceDir, "OriginalWorkbook.xlsx");
            string pfxPath = Path.Combine(sourceDir, "certificate.pfx");
            string pfxPassword = "yourPfxPassword";
            string comment = "XAdES Signed Document";

            // Ensure directories exist
            Directory.CreateDirectory(sourceDir);
            Directory.CreateDirectory(outputDir);

            // Create a sample workbook if it does not exist
            if (!File.Exists(sourceFile))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                wb.Save(sourceFile);
            }

            // Load or generate the certificate (PFX) as a byte array
            byte[] certData;
            if (File.Exists(pfxPath))
            {
                certData = File.ReadAllBytes(pfxPath);
            }
            else
            {
                using (RSA rsa = RSA.Create(2048))
                {
                    var req = new CertificateRequest("cn=AsposeDemo", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                    var cert = req.CreateSelfSigned(DateTimeOffset.Now.AddDays(-1), DateTimeOffset.Now.AddYears(1));
                    certData = cert.Export(X509ContentType.Pfx, pfxPassword);
                    File.WriteAllBytes(pfxPath, certData);
                }
            }

            // Load the workbook to be signed
            Workbook workbook = new Workbook(sourceFile);

            // Create a digital signature using the certificate data, password, comment, and current time
            DigitalSignature signature = new DigitalSignature(certData, pfxPassword, comment, DateTime.Now);

            // Enable XAdES (Basic XAdES) for the signature
            signature.XAdESType = XAdESType.XAdES;

            // Create a collection and add the signature to it
            DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
            signatureCollection.Add(signature);

            // Attach the digital signature collection to the workbook
            workbook.SetDigitalSignature(signatureCollection);

            // Save the signed workbook
            string outputFile = Path.Combine(outputDir, "SignedWorkbook.xlsx");
            workbook.Save(outputFile);

            // Verify the signature after saving
            Workbook signedWorkbook = new Workbook(outputFile);
            DigitalSignatureCollection loadedSignatures = signedWorkbook.GetDigitalSignature();

            foreach (DigitalSignature ds in loadedSignatures)
            {
                Console.WriteLine($"Comments: {ds.Comments}");
                Console.WriteLine($"Sign Time: {ds.SignTime}");
                Console.WriteLine($"Is Valid: {ds.IsValid}");
                Console.WriteLine($"XAdES Type: {ds.XAdESType}");
            }
        }
    }
}