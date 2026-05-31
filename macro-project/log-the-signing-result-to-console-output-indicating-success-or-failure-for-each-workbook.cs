using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using System.Security.Cryptography.X509Certificates;

namespace AsposeCellsDigitalSignatureDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths of workbooks to be signed
            string[] workbookPaths = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                // add more paths as needed
            };

            // Digital certificate used for signing
            string certificatePath = "certificate.pfx";
            string certificatePassword = "password";

            // Load the certificate once, after verifying the file exists
            X509Certificate2 certificate;
            try
            {
                if (!File.Exists(certificatePath))
                {
                    Console.WriteLine($"Certificate file not found: {certificatePath}");
                    return;
                }

                certificate = new X509Certificate2(certificatePath, certificatePassword);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load certificate: {ex.Message}");
                return;
            }

            foreach (string inputPath in workbookPaths)
            {
                try
                {
                    // Verify workbook file exists before loading
                    if (!File.Exists(inputPath))
                    {
                        Console.WriteLine($"Workbook file not found: {inputPath}");
                        continue;
                    }

                    // Load the workbook
                    using (Workbook workbook = new Workbook(inputPath))
                    {
                        // Create a digital signature collection and add a signature
                        DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                        DigitalSignature signature = new DigitalSignature(certificate, "Signed by Aspose.Cells", DateTime.Now);
                        signatures.Add(signature);

                        // Apply the digital signature to the workbook
                        workbook.SetDigitalSignature(signatures);

                        // Save the signed workbook (appending "_signed" to the original name)
                        string outputPath = Path.Combine(
                            Path.GetDirectoryName(inputPath) ?? string.Empty,
                            Path.GetFileNameWithoutExtension(inputPath) + "_signed.xlsx");

                        workbook.Save(outputPath);

                        // Verify signing result
                        bool isSigned = workbook.IsDigitallySigned;
                        Console.WriteLine($"Workbook '{inputPath}' signing {(isSigned ? "succeeded" : "failed")}.");
                    }
                }
                catch (Exception ex)
                {
                    // Log any failure during the signing process
                    Console.WriteLine($"Workbook '{inputPath}' signing failed: {ex.Message}");
                }
            }
        }
    }
}