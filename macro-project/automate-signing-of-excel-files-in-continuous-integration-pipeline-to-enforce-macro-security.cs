// Title: Automate digital signing of Excel workbooks with a PFX certificate in a CI/CD pipeline using Aspose.Cells for .NET
// AI Prompts: Create a .NET console application that accepts input.xlsx, output.xlsx, cert.pfx, and password arguments, loads the workbook with Aspose.Cells, applies a digital signature via SignatureCollection.Add, and writes the signed file. | Extend the signing tool to process a list of Excel files from a text file, applying the same PFX certificate to each workbook in a batch operation. | Add runtime detection for the SignatureCollection.Add method, log a warning if the API is missing, and let the CI build continue without failing.
// Common Searches: how to add a digital signature to an .xlsx file in a Jenkins pipeline using Aspose.Cells | C# console app for signing Excel workbooks with a PFX certificate | Aspose.Cells SignatureCollection.Add method example for CI/CD | fallback strategy when Aspose.Cells digital signature API is missing in a .NET project
// Tags: Aspose.Cells digital signature for Excel workbook | CI/CD pipeline Excel signing with PFX certificate | SignatureCollection.Add method example in .NET | automated macro security enforcement using Aspose.Cells | graceful degradation when signature API unavailable

using System;
using System.IO;
using Aspose.Cells;

// A .NET console utility that loads an Excel workbook, attempts to add a digital signature from a PFX certificate via Aspose.Cells' SignatureCollection, gracefully skips signing if the API is unavailable, and saves the (potentially signed) file—designed for integration into CI/CD pipelines to enforce macro security.
class ExcelSigner
{
    static void Main(string[] args)
    {
        // Expect four arguments: input file, output file, certificate file (PFX) and its password
        if (args.Length != 4)
        {
            Console.WriteLine("Usage: ExcelSigner <input.xlsx> <output.xlsx> <cert.pfx> <certPassword>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];
        string certPath = args[2];
        string certPassword = args[3];

        // Verify that required files exist
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        if (!File.Exists(certPath))
        {
            Console.WriteLine($"Certificate file not found: {certPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Add digital signature using dynamic invocation to avoid compile‑time binding issues
            try
            {
                dynamic wbDynamic = workbook;
                wbDynamic.SignatureCollection.Add(certPath, certPassword, "Signed by CI pipeline", "CI Server");
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                Console.WriteLine("Digital signature API not available in the current Aspose.Cells version. Skipping signing.");
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the (signed) workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
