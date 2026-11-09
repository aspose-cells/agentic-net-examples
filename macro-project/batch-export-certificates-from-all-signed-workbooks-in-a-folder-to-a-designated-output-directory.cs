// Title: How to batch export X509 certificates from digitally signed Excel workbooks using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells extract certificates from signed .xlsx files in a folder | batch export digital signature certificates from multiple Excel workbooks using Aspose.Cells | save X509 certificate from Excel workbook digital signature to .cer file with C# | process all signed Excel files in a directory and export their certificates using Aspose.Cells | use reflection to access DigitalSignatureCollection in Aspose.Cells for certificate extraction
// Tags: Aspose.Cells export certificate from DigitalSignatureCollection | batch process signed Excel workbooks C# | extract X509 certificate from .xlsx digital signature | save certificate as .cer file using Aspose.Cells | reflection based digital signature access Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The example scans a given input directory for .xlsx files, loads each workbook with Aspose.Cells, uses reflection to obtain the DigitalSignatureCollection, iterates through each signature to retrieve its X509 certificate, exports the certificate in DER (.cer) format, and writes the certificates to a specified output folder with filenames that include the workbook name and signature index.
class BatchExportCertificates
{
    static void Main(string[] args)
    {
        // Input folder containing signed workbooks
        string inputFolder = @"C:\InputWorkbooks";

        // Output folder where certificates will be saved
        string outputFolder = @"C:\ExportedCertificates";

        // Ensure the input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Process each Excel file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Safety check: skip if the file does not exist
            if (!File.Exists(filePath))
                continue;

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Use reflection to obtain the DigitalSignatureCollection (avoids compile‑time dependency)
                PropertyInfo sigCollectionProp = workbook.GetType().GetProperty("DigitalSignatureCollection");
                if (sigCollectionProp == null)
                {
                    Console.WriteLine($"No digital signature support in workbook: {filePath}");
                    continue;
                }

                object sigCollection = sigCollectionProp.GetValue(workbook);
                if (sigCollection == null)
                    continue;

                // Get the Count property of the collection
                PropertyInfo countProp = sigCollection.GetType().GetProperty("Count");
                int count = countProp != null ? (int)countProp.GetValue(sigCollection) : 0;
                if (count == 0)
                    continue;

                // Iterate through each signature
                for (int i = 0; i < count; i++)
                {
                    // Retrieve the signature via the indexer (Item property)
                    PropertyInfo itemProp = sigCollection.GetType().GetProperty("Item");
                    object signature = itemProp?.GetValue(sigCollection, new object[] { i });
                    if (signature == null)
                        continue;

                    // Get the attached certificate
                    PropertyInfo certProp = signature.GetType().GetProperty("Certificate");
                    object certObj = certProp?.GetValue(signature);
                    if (certObj == null)
                        continue;

                    // Export the certificate to DER format (.cer)
                    MethodInfo exportMethod = certObj.GetType().GetMethod("Export", new[] { typeof(X509ContentType) });
                    byte[] certBytes = exportMethod?.Invoke(certObj, new object[] { X509ContentType.Cert }) as byte[];
                    if (certBytes == null || certBytes.Length == 0)
                        continue;

                    // Build a unique file name: OriginalWorkbookName_SignatureIndex.cer
                    string workbookName = Path.GetFileNameWithoutExtension(filePath);
                    string certFileName = $"{workbookName}_Signature{i + 1}.cer";
                    string certFilePath = Path.Combine(outputFolder, certFileName);

                    // Write the certificate to the output file
                    File.WriteAllBytes(certFilePath, certBytes);
                }
            }
            catch (Exception ex)
            {
                // Log the error and continue with the next file
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Certificate export completed.");
    }
}
