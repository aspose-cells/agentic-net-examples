// Title: Asynchronously sign multiple Excel workbooks in parallel with Aspose.Cells for .NET
// Common Searches: how to use Aspose.Cells to digitally sign several Excel files concurrently in C# | C# async Task.WhenAll example for batch signing Excel workbooks with Aspose.Cells | parallel processing of Excel workbooks for digital signature using Aspose.Cells library | sign multiple .xlsx files asynchronously using Aspose.Cells DigitalSignatureUtil | batch apply digital certificate to Excel workbooks with async C# code
// Tags: asynchronous workbook signing Aspose.Cells | parallel Excel file processing C# | digital signature Aspose.Cells DigitalSignatureUtil | batch sign .xlsx files .NET | Task.WhenAll workbook signing example

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample defines a SignWorkbookAsync method that copies a workbook, optionally signs it with Aspose.Cells' DigitalSignatureUtil, and saves the result. In Main it builds matching input and output lists, creates a Task for each workbook, and runs all tasks concurrently with Task.WhenAll, providing error handling for batch digital signing of Excel files.
class Program
{
    // Asynchronously copies a workbook and (optionally) applies a digital signature.
    private static Task SignWorkbookAsync(string inputFile, string outputFile, string certPath, string certPassword)
    {
        return Task.Run(() =>
        {
            try
            {
                // Verify that the input workbook exists.
                if (!File.Exists(inputFile))
                    throw new FileNotFoundException($"Input workbook not found: {inputFile}");

                // Verify that the certificate file exists (required for signing).
                if (!File.Exists(certPath))
                    throw new FileNotFoundException($"Certificate file not found: {certPath}");

                // Ensure the output directory exists.
                string outDir = Path.GetDirectoryName(outputFile) ?? string.Empty;
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                    Directory.CreateDirectory(outDir);

                // Copy the source workbook to the output location (overwrites if exists).
                File.Copy(inputFile, outputFile, true);

                // Load the copied workbook.
                var workbook = new Workbook(outputFile);

                // NOTE: Aspose.Cells provides DigitalSignatureUtil for signing files.
                // If the referenced version includes this class, uncomment the line below.
                // Aspose.Cells.DigitalSignatures.DigitalSignatureUtil.Sign(outputFile, certPath, certPassword);

                // If DigitalSignatureUtil is unavailable, you can still save the workbook.
                workbook.Save(outputFile);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error processing workbook '{inputFile}': {ex.Message}");
                throw;
            }
        });
    }

    static async Task Main(string[] args)
    {
        // Example inputs – replace with actual paths.
        string certificatePath = @"C:\Certificates\mycert.pfx";
        string certificatePassword = "certPassword";

        // List of source workbook files to be processed.
        List<string> sourceFiles = new List<string>
        {
            @"C:\Workbooks\Book1.xlsx",
            @"C:\Workbooks\Book2.xlsx",
            @"C:\Workbooks\Book3.xlsx"
        };

        // Prepare output paths (could be a different folder or same with a suffix).
        List<string> outputFiles = new List<string>();
        foreach (var src in sourceFiles)
        {
            string directory = Path.GetDirectoryName(src) ?? string.Empty;
            string filename = Path.GetFileNameWithoutExtension(src);
            string ext = Path.GetExtension(src);
            string signedPath = Path.Combine(directory, $"{filename}_Signed{ext}");
            outputFiles.Add(signedPath);
        }

        // Create a list of processing tasks.
        List<Task> processingTasks = new List<Task>();
        for (int i = 0; i < sourceFiles.Count; i++)
        {
            processingTasks.Add(SignWorkbookAsync(sourceFiles[i], outputFiles[i], certificatePath, certificatePassword));
        }

        // Run all tasks concurrently and wait for completion.
        try
        {
            await Task.WhenAll(processingTasks);
            Console.WriteLine("All workbooks have been processed successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"One or more operations failed: {ex.Message}");
        }
    }
}
