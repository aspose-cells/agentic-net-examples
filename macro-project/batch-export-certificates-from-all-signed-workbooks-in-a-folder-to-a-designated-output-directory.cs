using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using System.Security.Cryptography.X509Certificates;

class BatchExportCertificates
{
    static void Main()
    {
        // Folder containing the workbooks to process
        string sourceFolder = @"C:\InputWorkbooks";

        // Folder where extracted certificates will be saved
        string outputFolder = @"C:\ExportedCertificates";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Supported Excel file extensions
        string[] extensions = new[] { ".xlsx", ".xlsm", ".xlsb", ".xls" };

        // Iterate through each file in the source folder
        foreach (string filePath in Directory.GetFiles(sourceFolder))
        {
            // Process only Excel files
            if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                continue;

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Check if the workbook is digitally signed
            if (!workbook.IsDigitallySigned)
                continue; // Skip unsigned workbooks

            // Retrieve the digital signature collection
            DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

            if (signatures == null)
                continue; // No signatures found

            int certIndex = 0;
            // Export each certificate found in the signature collection
            foreach (DigitalSignature signature in signatures)
            {
                // The DigitalSignature class provides access to the X509Certificate2 used for signing.
                // Export the certificate in DER format (.cer)
                X509Certificate2 cert = signature.Certificate;
                if (cert == null)
                    continue; // Skip if certificate data is unavailable

                byte[] certData = cert.Export(X509ContentType.Cert);

                // Build a unique file name for each exported certificate
                string outputFileName = $"{Path.GetFileNameWithoutExtension(filePath)}_cert_{certIndex}.cer";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Write the certificate bytes to disk
                File.WriteAllBytes(outputPath, certData);

                certIndex++;
            }

            // Dispose the workbook to release resources
            workbook.Dispose();
        }

        Console.WriteLine("Certificate export completed.");
    }
}