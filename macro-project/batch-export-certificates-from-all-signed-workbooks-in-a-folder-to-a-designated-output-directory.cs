using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using System.Security.Cryptography.X509Certificates;

namespace BatchExportCertificates
{
    class Program
    {
        static void Main()
        {
            // Folder containing the workbooks to process
            string sourceFolder = @"C:\InputWorkbooks";

            // Folder where extracted certificates will be saved
            string outputFolder = @"C:\ExportedCertificates";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            try
            {
                // Process each Excel file in the source folder
                foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
                {
                    // Verify the file exists before attempting to load
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(filePath);

                        // Skip if the workbook is not digitally signed
                        if (!workbook.IsDigitallySigned)
                        {
                            Console.WriteLine($"Skipping unsigned workbook: {Path.GetFileName(filePath)}");
                            continue;
                        }

                        // Retrieve the digital signature collection
                        DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                        if (signatures == null)
                        {
                            Console.WriteLine($"No signatures collection in: {Path.GetFileName(filePath)}");
                            continue;
                        }

                        int certIndex = 0;

                        // Iterate through each signature and export its certificate
                        foreach (DigitalSignature signature in signatures)
                        {
                            X509Certificate2 cert = signature.Certificate;

                            if (cert == null)
                            {
                                Console.WriteLine($"Signature {certIndex} in {Path.GetFileName(filePath)} does not contain a certificate.");
                                certIndex++;
                                continue;
                            }

                            // Export the certificate as a .cer file (DER encoded)
                            byte[] certData = cert.Export(X509ContentType.Cert);

                            // Build a unique file name for each exported certificate
                            string certFileName = $"{Path.GetFileNameWithoutExtension(filePath)}_cert{certIndex}.cer";
                            string certPath = Path.Combine(outputFolder, certFileName);

                            // Write the certificate bytes to disk
                            File.WriteAllBytes(certPath, certData);

                            Console.WriteLine($"Exported certificate to: {certPath}");

                            certIndex++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch export completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}