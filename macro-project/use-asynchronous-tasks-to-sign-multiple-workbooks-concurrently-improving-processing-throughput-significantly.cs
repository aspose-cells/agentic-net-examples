// Title: Async concurrent digital signing of multiple Excel workbooks with Aspose.Cells (C#)
// Description: This example demonstrates how to load an X509Certificate2 once and apply a digital signature to a collection of Excel files using Aspose.Cells. Each workbook is processed in its own async task, and Task.WhenAll is used to run the signatures in parallel, dramatically increasing throughput while handling errors gracefully.
// Keywords: Aspose.Cells | C# async digital signature | parallel workbook signing | Task.WhenAll | X509Certificate2 reuse | Excel batch signing | digital signature .NET | concurrent Excel processing
// Common Searches: sign multiple Excel files concurrently C# | Aspose.Cells async digital signature example | parallel workbook signing with Task.WhenAll | reuse X509Certificate2 for batch signing | how to improve Excel signing throughput .NET
// Developer Intent: Implement high‑throughput batch signing of Excel workbooks by leveraging asynchronous tasks and a single certificate instance.
// Use Cases: Mass signing of financial statements before external distribution. | Automated background service that signs thousands of generated invoices in parallel. | Secure archiving of regulatory reports where each file must carry a trusted timestamp.
// AI Prompts: Create a version of the code that limits parallelism with a SemaphoreSlim and a configurable max degree of concurrency. | Add logging that records successful and failed signatures to a structured JSON file for audit purposes. | Refactor the sample to support cancellation via a CancellationToken and report progress through IProgress<T>.

using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

// This example demonstrates how to load an X509Certificate2 once and apply a digital signature to a collection of Excel files using Aspose.Cells. Each workbook is processed in its own async task, and Task.WhenAll is used to run the signatures in parallel, dramatically increasing throughput while handling errors gracefully.
public class WorkbookSigner
{
    // Asynchronously signs a single workbook and saves the signed copy.
    private static async Task SignWorkbookAsync(string sourcePath, string destinationPath, X509Certificate2 certificate)
    {
        try
        {
            // Verify source workbook exists.
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

            // Load the workbook from the source file.
            using (Workbook workbook = new Workbook(sourcePath))
            {
                // Create a digital signature collection and add a signature.
                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                DigitalSignature signature = new DigitalSignature(certificate, "Automated Signature", DateTime.Now);
                signatures.Add(signature);

                // Add the digital signature to the workbook.
                workbook.AddDigitalSignature(signatures);

                // Ensure destination directory exists.
                string destDir = Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                    Directory.CreateDirectory(destDir);

                // Save the signed workbook to the destination path.
                workbook.Save(destinationPath);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error signing workbook '{sourcePath}': {ex.Message}");
        }

        // Simulate asynchronous work (optional).
        await Task.Yield();
    }

    // Signs multiple workbooks concurrently using asynchronous tasks.
    public static async Task SignWorkbooksConcurrentlyAsync(IEnumerable<(string source, string destination)> files, string certPath, string certPassword)
    {
        try
        {
            // Verify certificate file exists.
            if (!File.Exists(certPath))
                throw new FileNotFoundException($"Certificate file not found: {certPath}");

            // Load the certificate once; it will be reused for all workbooks.
            using (X509Certificate2 certificate = new X509Certificate2(certPath, certPassword))
            {
                List<Task> signingTasks = new List<Task>();

                foreach (var (source, destination) in files)
                {
                    // Start a signing task for each workbook.
                    signingTasks.Add(SignWorkbookAsync(source, destination, certificate));
                }

                // Await all signing tasks to complete.
                await Task.WhenAll(signingTasks);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error during signing process: {ex.Message}");
        }
    }

    // Example usage.
    public static async Task Main()
    {
        try
        {
            // Define source and destination file pairs.
            var filesToSign = new List<(string source, string destination)>
            {
                (@"C:\Docs\Report1.xlsx", @"C:\Signed\Report1_Signed.xlsx"),
                (@"C:\Docs\Report2.xlsx", @"C:\Signed\Report2_Signed.xlsx"),
                (@"C:\Docs\Report3.xlsx", @"C:\Signed\Report3_Signed.xlsx")
            };

            // Path to the PFX certificate and its password.
            string certificatePath = @"C:\Certificates\mycert.pfx";
            string certificatePassword = "yourPassword";

            // Sign all workbooks concurrently.
            await SignWorkbooksConcurrentlyAsync(filesToSign, certificatePath, certificatePassword);

            Console.WriteLine("All workbooks have been processed.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
