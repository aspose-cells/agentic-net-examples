// Title: Batch Export X509 Certificates from Digitally Signed Excel Workbooks with Aspose.Cells for .NET (C#)
// Description: Scans a folder for Excel files, loads each workbook with Aspose.Cells, detects digital signatures, extracts the embedded X509Certificate2 objects, and writes them as DER‑encoded .cer files to a target directory while handling errors per file.
// Keywords: Aspose.Cells | C# | .NET | digital signature | certificate extraction | X509 certificate | batch export | Excel workbook | export .cer file | automation | compliance audit | signed Excel files
// Common Searches: extract certificates from signed Excel workbooks C# | Aspose.Cells batch export digital signature certificates | how to save X509 certificates from Excel files | C# code to export .cer files from digitally signed workbooks | automate certificate extraction from multiple Excel files
// Developer Intent: Programmatically retrieve every X509 certificate embedded in digitally signed Excel workbooks within a folder and save each as a .cer file in a specified output location.
// Use Cases: Create a repository of signing certificates for regulatory compliance across hundreds of reports. | Prepare .cer files before migrating signed workbooks to a new document management system. | Integrate certificate extraction into a CI/CD pipeline to verify that generated workbooks are signed by authorized keys.
// AI Prompts: Generate C# code that recursively scans subfolders for signed Excel files and exports each signature's X509 certificate as a .cer file using Aspose.Cells. | Add logging to the batch exporter that records the source workbook path, signature index, and output certificate file name. | Enhance the example to return a summary object containing total workbooks processed, signatures found, and certificates successfully exported.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using System.Security.Cryptography.X509Certificates;

namespace AsposeCellsExamples
{
    // Scans a folder for Excel files, loads each workbook with Aspose.Cells, detects digital signatures, extracts the embedded X509Certificate2 objects, and writes them as DER‑encoded .cer files to a target directory while handling errors per file.
    public class CertificateBatchExporter
    {
        // Exports certificates from all digitally signed workbooks in a folder.
        public static void ExportCertificates(string sourceFolder, string outputFolder)
        {
            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the source folder (non‑recursive)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                // Filter supported Excel extensions
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xlsm" && ext != ".xlsb" && ext != ".xls")
                    continue;

                // Verify the file exists before loading
                if (!File.Exists(filePath))
                    continue;

                try
                {
                    // Load the workbook
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        // Skip if the workbook is not digitally signed
                        if (!workbook.IsDigitallySigned)
                            continue;

                        // Retrieve the digital signature collection
                        DigitalSignatureCollection signatures = workbook.GetDigitalSignature();
                        if (signatures == null)
                            continue;

                        int sigIndex = 0;
                        foreach (DigitalSignature signature in signatures)
                        {
                            // Attempt to obtain the X509 certificate from the signature via reflection
                            X509Certificate2 cert = null;
                            try
                            {
                                var certProp = typeof(DigitalSignature).GetProperty("Certificate");
                                if (certProp != null)
                                {
                                    cert = certProp.GetValue(signature) as X509Certificate2;
                                }
                            }
                            catch
                            {
                                // Ignore reflection errors
                            }

                            if (cert == null)
                                continue; // Unable to retrieve certificate

                            // Export the certificate in DER format (.cer)
                            byte[] certData = cert.Export(X509ContentType.Cert);
                            string fileName = $"{Path.GetFileNameWithoutExtension(filePath)}_cert_{sigIndex}.cer";
                            string outPath = Path.Combine(outputFolder, fileName);
                            File.WriteAllBytes(outPath, certData);
                            sigIndex++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log or handle errors for individual files without stopping the batch
                    Console.Error.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }
        }
    }

    public class Program
    {
        // Entry point required for compilation
        public static void Main(string[] args)
        {
            try
            {
                // Example folders – adjust paths as needed or pass via command line arguments
                string sourceFolder = args.Length > 0 ? args[0] : @"C:\InputWorkbooks";
                string outputFolder = args.Length > 1 ? args[1] : @"C:\ExportedCertificates";

                // Validate source folder existence
                if (!Directory.Exists(sourceFolder))
                {
                    Console.Error.WriteLine($"Source folder does not exist: {sourceFolder}");
                    return;
                }

                CertificateBatchExporter.ExportCertificates(sourceFolder, outputFolder);
                Console.WriteLine("Certificate export completed.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
