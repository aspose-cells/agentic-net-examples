// Title: Export digital signature certificates from multiple Excel workbooks asynchronously with Aspose.Cells for .NET
// AI Prompts: Write a C# async method that iterates over a collection of .xlsx file paths, loads each workbook with Aspose.Cells, extracts any digital‑signature certificates via reflection, and writes the certificate details (subject, issuer, validity dates, thumbprint) to a .cert.txt file in a given output directory. | Extend the asynchronous exporter to also capture each certificate's serial number and output the results in JSON format instead of plain text. | Add a configurable MaxDegreeOfParallelism argument to the export routine so that only a limited number of workbooks are processed concurrently while still using async tasks.
// Common Searches: how to extract digital signature certificates from Excel files using Aspose.Cells in C# | async processing of multiple .xlsx workbooks to export certificates .NET | using reflection to access DigitalSignatureCollection in Aspose.Cells | parallel export of X509Certificate2 information from Excel workbooks | limit parallel tasks when exporting certificates from Excel files with Aspose.Cells
// Tags: asynchronous export of Excel digital signatures Aspose.Cells | reflection access DigitalSignatureCollection C# | write X509Certificate2 details to text file | parallel processing of .xlsx workbooks | configurable degree of parallelism certificate extraction

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using System.Security.Cryptography.X509Certificates;
using System.Reflection;

// The example defines a CertificateExporter class that asynchronously processes a list of Excel workbook paths. For each .xlsx file it loads the workbook with Aspose.Cells, uses reflection to obtain the DigitalSignatureCollection and each signature's Certificate, and writes key certificate properties (subject, issuer, validity period, thumbprint) to a .cert.txt file in a specified output folder. The exporter runs each workbook export in its own Task and awaits all tasks, enabling high‑performance processing of large file sets.
public class CertificateExporter
{
    // Export certificates from multiple workbooks asynchronously.
    public static async Task ExportCertificatesAsync(IEnumerable<string> workbookPaths, string outputFolder)
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        var exportTasks = new List<Task>();

        foreach (var path in workbookPaths)
        {
            exportTasks.Add(Task.Run(() =>
            {
                try
                {
                    // Verify the workbook file exists.
                    if (!File.Exists(path))
                    {
                        Console.WriteLine($"File not found: {path}");
                        return;
                    }

                    // Load the workbook.
                    var workbook = new Workbook(path);

                    // Prepare the output file name.
                    var outputFile = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(path) + ".cert.txt");

                    using (var writer = new StreamWriter(outputFile))
                    {
                        // Try to obtain the DigitalSignatureCollection via reflection
                        // (the property may not exist in older Aspose.Cells versions).
                        PropertyInfo sigProp = workbook.GetType().GetProperty("DigitalSignatureCollection",
                            BindingFlags.Public | BindingFlags.Instance);

                        if (sigProp != null)
                        {
                            var signaturesObj = sigProp.GetValue(workbook);
                            if (signaturesObj is System.Collections.IEnumerable signatures && signatures != null)
                            {
                                int index = 0;
                                foreach (var sig in signatures)
                                {
                                    index++;
                                    try
                                    {
                                        // Each signature is expected to be of type DigitalSignature.
                                        // Use reflection to get the Certificate property.
                                        PropertyInfo certProp = sig.GetType().GetProperty("Certificate",
                                            BindingFlags.Public | BindingFlags.Instance);
                                        if (certProp == null)
                                        {
                                            writer.WriteLine($"Signature #{index}: Certificate property not found.");
                                            continue;
                                        }

                                        var certObj = certProp.GetValue(sig);
                                        if (certObj == null)
                                        {
                                            writer.WriteLine($"Signature #{index}: Certificate is null.");
                                            continue;
                                        }

                                        // Convert to X509Certificate2.
                                        X509Certificate2 cert = certObj as X509Certificate2 ??
                                            new X509Certificate2(certObj as byte[] ?? new byte[0]);

                                        writer.WriteLine($"Certificate #{index}");
                                        writer.WriteLine($"Subject: {cert.Subject}");
                                        writer.WriteLine($"Issuer: {cert.Issuer}");
                                        writer.WriteLine($"Valid From: {cert.NotBefore}");
                                        writer.WriteLine($"Valid To: {cert.NotAfter}");
                                        writer.WriteLine($"Thumbprint: {cert.GetCertHashString()}");
                                        writer.WriteLine(new string('-', 40));
                                    }
                                    catch (Exception sigEx)
                                    {
                                        Console.WriteLine($"Error processing signature #{index} in '{path}': {sigEx.Message}");
                                    }
                                }

                                // If no signatures were enumerated.
                                if (index == 0)
                                {
                                    writer.WriteLine("No digital signatures found in this workbook.");
                                }
                            }
                            else
                            {
                                writer.WriteLine("DigitalSignatureCollection is empty or not enumerable.");
                            }
                        }
                        else
                        {
                            writer.WriteLine("Digital signatures are not supported in this version of Aspose.Cells.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{path}': {ex.Message}");
                }
            }));
        }

        // Await all export tasks to complete.
        await Task.WhenAll(exportTasks);
    }
}

public class Program
{
    // Entry point for the console application.
    public static async Task Main(string[] args)
    {
        try
        {
            string inputFolder = args.Length > 0 ? args[0] : @"C:\Workbooks";
            string outputFolder = args.Length > 1 ? args[1] : @"C:\CertificatesExport";

            // Verify input folder exists.
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Get all .xlsx files from the input folder.
            var files = Directory.GetFiles(inputFolder, "*.xlsx");

            await CertificateExporter.ExportCertificatesAsync(files, outputFolder);
            Console.WriteLine("Certificate export completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
