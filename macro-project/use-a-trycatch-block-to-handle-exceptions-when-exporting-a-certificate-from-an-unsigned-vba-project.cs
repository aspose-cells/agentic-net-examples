// Title: Export VBA Project Certificate with Try‑Catch Error Handling in Aspose.Cells for .NET
// Description: Demonstrates how to load an XLSM workbook, verify the presence of a VBA project, extract its certificate (CertRawData) and write it to a .cer file while using try‑catch blocks to handle missing files, unsigned projects, I/O failures, and optional workbook saving.
// Keywords: Aspose.Cells export VBA certificate | C# VBA project CertRawData | handle unsigned VBA project | try catch Aspose.Cells | export .cer from XLSM | VbaProject error handling | save XLSM workbook C#
// Common Searches: export VBA certificate Aspose.Cells C# | how to catch errors when extracting VBA CertRawData | C# code to detect unsigned VBA project in XLSM | exception handling for loading and saving XLSM with VBA | Aspose.Cells VBA project certificate export example
// Developer Intent: The developer needs a reliable way to export a VBA project's signing certificate and gracefully handle scenarios such as missing files, unsigned projects, or I/O errors.
// Use Cases: Export a signed VBA project's certificate to a .cer file and log a clear message if the project is unsigned. | Load an XLSM workbook only after confirming the file exists, preventing FileNotFoundException. | Save the processed workbook in Xlsm format while capturing any save‑time failures.
// AI Prompts: Write a C# method using Aspose.Cells that extracts a VBA project's certificate, returns a boolean success flag, and includes detailed try‑catch handling for missing files, unsigned projects, and I/O errors. | Create unit tests for ExportVbaCertificateDemo covering: missing input file, unsigned VBA project, and successful certificate export. | Refactor the sample to separate certificate extraction and workbook saving into reusable helper functions with proper exception propagation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Demonstrates how to load an XLSM workbook, verify the presence of a VBA project, extract its certificate (CertRawData) and write it to a .cer file while using try‑catch blocks to handle missing files, unsigned projects, I/O failures, and optional workbook saving.
public class ExportVbaCertificateDemo
{
    public static void Run()
    {
        const string inputFile = "UnsignedVba.xlsm";
        const string outputCertFile = "ExportedVbaCertificate.cer";
        const string outputWorkbookFile = "UnsignedVba_Processed.xlsm";

        Workbook workbook = null;

        // Load the workbook safely
        try
        {
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"Input file '{inputFile}' was not found.");

            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Access the VBA project (may be null if none exists)
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
                throw new InvalidOperationException("The workbook does not contain a VBA project.");

            // Retrieve the certificate raw data
            byte[] certData = vbaProject.CertRawData;

            // If the project is unsigned, CertRawData will be null or empty – treat this as an error
            if (certData == null || certData.Length == 0)
                throw new InvalidOperationException("VBA project is not signed; certificate data is unavailable.");

            // Export the certificate to a .cer file
            File.WriteAllBytes(outputCertFile, certData);
            Console.WriteLine("Certificate exported successfully.");
        }
        catch (Exception ex)
        {
            // Handle all exceptions (including unsigned project case)
            Console.WriteLine($"Error exporting certificate: {ex.Message}");
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        try
        {
            workbook.Save(outputWorkbookFile, SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            ExportVbaCertificateDemo.Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
