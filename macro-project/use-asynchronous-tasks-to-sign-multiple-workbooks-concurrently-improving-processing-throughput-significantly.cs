using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsConcurrentSigning
{
    class Program
    {
        // Asynchronously signs a single workbook and saves the signed copy.
        private static async Task SignWorkbookAsync(string inputPath, string outputPath, X509Certificate2 certificate)
        {
            try
            {
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input workbook not found: {inputPath}");

                // Load the workbook.
                using (Workbook workbook = new Workbook(inputPath))
                {
                    // Create a digital signature.
                    DigitalSignature signature = new DigitalSignature(certificate, "Concurrent Signature", DateTime.Now);

                    // Add the signature to a collection and attach it to the workbook.
                    DigitalSignatureCollection signatures = new DigitalSignatureCollection { signature };
                    workbook.AddDigitalSignature(signatures);

                    // Save the signed workbook.
                    workbook.Save(outputPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error signing workbook '{inputPath}': {ex.Message}");
            }

            // Simulate asynchronous work (optional).
            await Task.CompletedTask;
        }

        static async Task Main(string[] args)
        {
            try
            {
                // Paths of workbooks to be signed.
                List<string> inputFiles = new List<string>
                {
                    "Workbook1.xlsx",
                    "Workbook2.xlsx",
                    "Workbook3.xlsx"
                };

                // Corresponding output paths for signed workbooks.
                List<string> outputFiles = new List<string>
                {
                    "Workbook1_Signed.xlsx",
                    "Workbook2_Signed.xlsx",
                    "Workbook3_Signed.xlsx"
                };

                // Verify certificate file exists before loading.
                string certPath = "certificate.pfx";
                string certPassword = "password";

                if (!File.Exists(certPath))
                    throw new FileNotFoundException($"Certificate file not found: {certPath}");

                // Load the signing certificate.
                X509Certificate2 cert = new X509Certificate2(certPath, certPassword, X509KeyStorageFlags.MachineKeySet);

                // Create a list to hold signing tasks.
                List<Task> signingTasks = new List<Task>();

                // Launch signing tasks concurrently.
                for (int i = 0; i < inputFiles.Count; i++)
                {
                    string input = inputFiles[i];
                    string output = outputFiles[i];
                    signingTasks.Add(SignWorkbookAsync(input, output, cert));
                }

                // Wait for all signing operations to complete.
                await Task.WhenAll(signingTasks);

                Console.WriteLine("All workbooks have been signed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}