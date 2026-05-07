using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace XAdESSignatureProcessing
{
    public class XAdESProcessor
    {
        // Signs an incoming XLSX file with XAdES signature and verifies it.
        public static void Process(string sourceXlsxPath, string pfxPath, string pfxPassword, string outputSignedPath)
        {
            // Ensure source workbook exists
            if (!File.Exists(sourceXlsxPath))
                throw new FileNotFoundException($"Source workbook not found: {sourceXlsxPath}");

            // Ensure certificate exists; if not, create a self‑signed one for demo purposes
            if (!File.Exists(pfxPath))
                CreateSelfSignedCertificate(pfxPath, pfxPassword);

            // Load the workbook to be signed
            Workbook workbook = new Workbook(sourceXlsxPath);

            // Load the certificate data (PFX) as a byte array
            byte[] certData = File.ReadAllBytes(pfxPath);

            // Create a digital signature
            DigitalSignature signature = new DigitalSignature(certData, pfxPassword, "XAdES Signed Document", DateTime.UtcNow);
            signature.XAdESType = XAdESType.XAdES; // Basic XAdES

            // Add the signature to a collection
            DigitalSignatureCollection signatures = new DigitalSignatureCollection();
            signatures.Add(signature);

            // Attach the digital signature collection to the workbook
            workbook.SetDigitalSignature(signatures);

            // Save the signed workbook
            workbook.Save(outputSignedPath, SaveFormat.Xlsx);

            // ---- Verification Phase ----
            // Load the signed workbook
            Workbook signedWorkbook = new Workbook(outputSignedPath);

            // Retrieve the digital signatures
            DigitalSignatureCollection loadedSignatures = signedWorkbook.GetDigitalSignature();

            // Output signature details
            foreach (DigitalSignature ds in loadedSignatures)
            {
                Console.WriteLine($"Comments   : {ds.Comments}");
                Console.WriteLine($"Sign Time  : {ds.SignTime:u}");
                Console.WriteLine($"Is Valid   : {ds.IsValid}");
                Console.WriteLine($"XAdES Type : {ds.XAdESType}");
                Console.WriteLine(new string('-', 40));
            }
        }

        // Creates a temporary self‑signed certificate and saves it as a PFX file.
        private static void CreateSelfSignedCertificate(string pfxPath, string password)
        {
            using RSA rsa = RSA.Create(2048);
            var request = new CertificateRequest(
                "CN=AsposeDemoCertificate",
                rsa,
                HashAlgorithmName.SHA256,
                RSASignaturePadding.Pkcs1);

            // Basic constraints - self signed
            request.CertificateExtensions.Add(
                new X509BasicConstraintsExtension(true, false, 0, true));

            // Key usage
            request.CertificateExtensions.Add(
                new X509KeyUsageExtension(
                    X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment,
                    true));

            // Enhanced key usage
            request.CertificateExtensions.Add(
                new X509EnhancedKeyUsageExtension(
                    new OidCollection { new Oid("1.3.6.1.5.5.7.3.3") }, // Code signing
                    true));

            DateTimeOffset notBefore = DateTimeOffset.UtcNow.AddDays(-1);
            DateTimeOffset notAfter = notBefore.AddYears(2);
            using X509Certificate2 cert = request.CreateSelfSigned(notBefore, notAfter);

            // Export as PFX
            byte[] pfxBytes = cert.Export(X509ContentType.Pfx, password);
            File.WriteAllBytes(pfxPath, pfxBytes);
        }

        // Example usage
        public static void Main()
        {
            string sourceFile = "input.xlsx";               // Path to incoming XLSX
            string certificateFile = "signer.pfx";          // Path to PFX certificate
            string certificatePassword = "pfxPassword";     // Certificate password
            string signedOutput = "signed_output.xlsx";     // Destination for signed file

            Process(sourceFile, certificateFile, certificatePassword, signedOutput);
        }
    }
}