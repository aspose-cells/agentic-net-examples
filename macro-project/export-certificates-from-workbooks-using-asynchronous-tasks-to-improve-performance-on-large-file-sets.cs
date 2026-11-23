// Title: Asynchronously export digital certificates from multiple Excel .xlsx workbooks using Aspose.Cells for .NET
// Common Searches: how to extract digital signature certificates from Excel files with Aspose.Cells in C# | using Task.WhenAll to process many .xlsx files in parallel with Aspose.Cells | C# code to export X509 certificates from workbook digital signatures asynchronously | batch export of certificates from Excel workbooks using Aspose.Cells .NET library | parallel extraction of digital signatures from multiple Excel workbooks C# Aspose.Cells
// Tags: asynchronous certificate export Aspose.Cells | digital signature collection extraction .xlsx | parallel workbook processing C# | export X509Certificate2 from Excel workbook | batch certificate extraction Aspose.Cells .NET

using System;
using System.IO;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// The utility scans a folder for .xlsx files, loads each workbook with Aspose.Cells, accesses its DigitalSignatureCollection, and writes each embedded X509Certificate2 to a separate .cer file in an output directory. Each workbook is processed on its own Task, and Task.WhenAll ensures all asynchronous operations complete before the program exits.
class Program
{
    static async Task Main(string[] args)
    {
        // Input folder containing workbooks and output folder for certificates
        string inputFolder = args.Length > 0 ? args[0] : "InputWorkbooks";
        string outputFolder = args.Length > 1 ? args[1] : "ExportedCertificates";

        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the input folder
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        var tasks = new List<Task>();

        // Process each workbook in its own task
        foreach (string filePath in workbookFiles)
        {
            tasks.Add(Task.Run(() => ExportCertificates(filePath, outputFolder)));
        }

        // Wait for all tasks to complete
        await Task.WhenAll(tasks);
    }

    static void ExportCertificates(string workbookPath, string outputFolder)
    {
        try
        {
            // Ensure the workbook file exists before loading
            if (!File.Exists(workbookPath))
                return;

            // Load the workbook using dynamic to safely access DigitalSignatureCollection
            dynamic workbook = new Workbook(workbookPath);

            // Retrieve digital signatures; if the property does not exist, signatures will be null
            var signatures = workbook.DigitalSignatureCollection as IEnumerable<object>;

            if (signatures == null)
                return; // No signatures collection available

            string workbookName = Path.GetFileNameWithoutExtension(workbookPath);
            int certIndex = 1;

            // Iterate through each digital signature
            foreach (dynamic signature in signatures)
            {
                // Access the Certificate property dynamically
                X509Certificate2 cert = signature.Certificate as X509Certificate2;
                if (cert == null)
                    continue;

                // Build a unique file name for each exported certificate
                string certFileName = $"{workbookName}_cert_{certIndex}.cer";
                string certPath = Path.Combine(outputFolder, certFileName);

                // Save the certificate in DER format
                File.WriteAllBytes(certPath, cert.Export(X509ContentType.Cert));

                certIndex++;
            }
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed; for now we just write to console
            Console.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
        }
    }
}
