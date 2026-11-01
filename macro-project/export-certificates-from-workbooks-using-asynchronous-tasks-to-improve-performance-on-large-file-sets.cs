// Title: Export digital certificates from multiple Excel workbooks asynchronously with Aspose.Cells for .NET
// Common Searches: how to extract X509 certificates from Excel files using Aspose.Cells in C# | async processing of workbook digital signatures with Aspose.Cells | parallel export of .cer files from .xlsx workbooks in .NET | using reflection to read DigitalSignatures collection in Aspose.Cells | batch certificate extraction from a folder of Excel workbooks C#
// Tags: asynchronous workbook certificate extraction Aspose.Cells | digital signatures extraction .xlsx C# | parallel processing of Excel workbooks Aspose.Cells | export X509Certificate2 to .cer file C# | reflection access DigitalSignatures Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Aspose.Cells;

// The example defines an async method that scans a directory for .xlsx files, loads each workbook with Aspose.Cells, uses reflection to obtain its DigitalSignatures collection, extracts any X509Certificate2 objects, exports each certificate in DER format as a .cer file with a unique name, and runs the extraction for all workbooks in parallel using Task.Run and Task.WhenAll.
public class CertificateExporter
{
    // Export certificates from all workbook files in the input folder to the output folder.
    // Each certificate is saved as a .cer file named "<WorkbookName>_cert_<index>.cer".
    public static async Task ExportCertificatesAsync(string inputFolder, string outputFolder)
    {
        // Ensure output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files (you can adjust the search pattern as needed).
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories);

        // List to hold all running tasks.
        List<Task> exportTasks = new List<Task>();

        foreach (string workbookPath in workbookFiles)
        {
            // Start a task for each workbook to run in parallel.
            exportTasks.Add(Task.Run(() =>
            {
                try
                {
                    // Verify the file exists before loading.
                    if (!File.Exists(workbookPath))
                    {
                        Console.WriteLine($"File not found: {workbookPath}");
                        return;
                    }

                    // Load the workbook.
                    Workbook workbook = new Workbook(workbookPath);

                    // Use reflection/dynamic to access digital signatures (avoids compile‑time dependency on the exact API version).
                    var signaturesProp = workbook.GetType().GetProperty("DigitalSignatures");
                    if (signaturesProp == null)
                    {
                        Console.WriteLine($"No digital signatures property found in workbook '{workbookPath}'.");
                        return;
                    }

                    dynamic signatures = signaturesProp.GetValue(workbook);
                    if (signatures == null)
                    {
                        Console.WriteLine($"No digital signatures present in workbook '{workbookPath}'.");
                        return;
                    }

                    int count = signatures.Count;
                    for (int i = 0; i < count; i++)
                    {
                        try
                        {
                            dynamic signature = signatures[i];
                            X509Certificate2 cert = signature.Certificate as X509Certificate2;

                            if (cert != null)
                            {
                                // Export the certificate in DER format.
                                byte[] certBytes = cert.Export(X509ContentType.Cert);

                                // Build a unique file name for the certificate.
                                string certFileName = $"{Path.GetFileNameWithoutExtension(workbookPath)}_cert_{i + 1}.cer";
                                string certFilePath = Path.Combine(outputFolder, certFileName);

                                // Write the certificate to disk.
                                File.WriteAllBytes(certFilePath, certBytes);
                            }
                        }
                        catch (Exception innerEx)
                        {
                            Console.WriteLine($"Error processing signature {i + 1} in workbook '{workbookPath}': {innerEx.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing workbook '{workbookPath}': {ex.Message}");
                }
            }));
        }

        // Await completion of all export tasks.
        await Task.WhenAll(exportTasks);
    }

    // Example usage.
    public static async Task Main(string[] args)
    {
        // Input folder containing workbooks.
        string inputFolder = @"C:\Workbooks";

        // Output folder where certificates will be saved.
        string outputFolder = @"C:\Certificates";

        await ExportCertificatesAsync(inputFolder, outputFolder);

        Console.WriteLine("Certificate export completed.");
    }
}
