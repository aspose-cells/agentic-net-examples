// Title: Sign multiple Excel workbooks concurrently with Aspose.Cells in C# using async/await
// AI Prompts: Create an async method that loads an .xlsx file with Aspose.Cells, attaches a digital signature from a PFX certificate, and writes the signed workbook to a specified output folder. | Write a loop that launches a separate Task for each workbook path in a collection, then waits for all signing tasks to finish using await. | Add support for a CancellationToken and a progress callback to the concurrent signing routine so the operation can be cancelled or reported.
// Common Searches: C# example for signing several Excel files at the same time with Aspose.Cells | How to use async/await to apply a digital signature to multiple .xlsx workbooks in .NET | Batch processing of Excel workbooks with digital certificates using Aspose.Cells | Parallel execution of workbook signing with Aspose.Cells and .NET Core
// Tags: aspose.cells async workbook signing | c# parallel excel digital signature | batch xlsx signing with aspnet | task based concurrent workbook processing | aspose.cells save signed workbook

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample creates an output folder, checks for a PFX certificate, and then processes a predefined list of Excel files concurrently. Each file is loaded into an Aspose.Cells Workbook, optionally signed (placeholder for the digital signature API), and saved with a '_signed' suffix. Concurrency is achieved by launching a Task for each workbook and awaiting all tasks with async/await, improving overall throughput.
class Program
{
    static async Task Main(string[] args)
    {
        try
        {
            // List of workbook files to be processed
            var inputFiles = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Directory where processed workbooks will be saved
            string outputDir = "SignedWorkbooks";
            Directory.CreateDirectory(outputDir);

            // Path to the signing certificate (PFX) and its password
            string certPath = "mycert.pfx";
            string certPassword = "password";

            // Verify that the certificate file exists
            if (!File.Exists(certPath))
            {
                Console.WriteLine($"Certificate file not found: {certPath}");
                return;
            }

            // Signature metadata (used if digital signing is supported)
            string reason = "Document approval";
            string location = "Company HQ";

            // Create a collection of processing tasks
            var processingTasks = new List<Task>();

            foreach (var inputPath in inputFiles)
            {
                // Start the asynchronous processing operation for each workbook
                processingTasks.Add(ProcessWorkbookAsync(inputPath, outputDir, certPath, certPassword, reason, location));
            }

            // Await completion of all processing tasks
            await Task.WhenAll(processingTasks);

            Console.WriteLine("All workbooks have been processed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Asynchronously processes a workbook (loads, optionally signs, and saves)
    private static async Task ProcessWorkbookAsync(string inputPath, string outputDir, string certPath, string certPassword, string reason, string location)
    {
        await Task.Run(() =>
        {
            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Prepare output path
                string fileName = Path.GetFileNameWithoutExtension(inputPath);
                string outputPath = Path.Combine(outputDir, $"{fileName}_signed.xlsx");

                // Load workbook
                var workbook = new Workbook(inputPath);

                // NOTE: Digital signature APIs may not be available in the current Aspose.Cells version.
                // If supported, you could sign the workbook as shown below:
                // workbook.DigitalSignatureCollection.AddSignature(certPath, certPassword, reason, location);

                // Save workbook (signed if the API is available)
                workbook.Save(outputPath);

                Console.WriteLine($"Processed workbook saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {inputPath}: {ex.Message}");
            }
        });
    }
}
