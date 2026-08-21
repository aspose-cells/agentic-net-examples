// Title: Export a Signed VBA Project’s Digital Certificate to a .cer File with Aspose.Cells for .NET (C#)
// Description: Loads a macro‑enabled workbook, checks if its VBA project is signed, extracts the raw certificate bytes via VbaProject.CertRawData, and writes them to a .cer file while handling missing data and runtime errors.
// Keywords: Aspose.Cells | C# VBA certificate export | VbaProject CertRawData | extract VBA signing certificate | save .cer file | macro-enabled workbook | signed VBA project | Excel VBA certificate extraction
// Common Searches: export VBA certificate Aspose.Cells C# | how to get VBA project signing certificate from .xlsm using C# | save VBA digital certificate to .cer with Aspose | retrieve CertRawData from VbaProject | extract signed macro certificate .NET
// Developer Intent: Retrieve and store the digital certificate of a signed VBA project embedded in an Excel workbook.
// Use Cases: Verify the authenticity of a macro‑enabled workbook by exporting its signing certificate for external validation. | Archive VBA signing certificates to meet compliance or audit requirements across multiple workbooks. | Compare certificates from different workbooks to detect unauthorized macro modifications. | Integrate the exported .cer file into a PKI workflow for automated trust checks.
// AI Prompts: Generate C# code that loads an .xlsm file with Aspose.Cells, checks if the VBA project is signed, and writes the certificate to a specified .cer path. | Add robust error handling to the VBA certificate export routine, covering scenarios such as missing workbook, unsigned VBA project, and empty certificate data. | Show how to load the exported .cer file into an X509Certificate2 object and validate its thumbprint against a trusted store. | Create a batch script that processes a folder of .xlsm files, exporting each signed VBA project's certificate to a corresponding .cer file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a macro‑enabled workbook, checks if its VBA project is signed, extracts the raw certificate bytes via VbaProject.CertRawData, and writes them to a .cer file while handling missing data and runtime errors.
    public class ExportVbaCertificate
    {
        public static void Run()
        {
            try
            {
                // Path to the workbook that contains a signed VBA project
                string workbookPath = "SignedWorkbook.xlsm";

                // Path where the extracted certificate will be saved
                string certificateOutputPath = "VbaCertificate.cer";

                // Verify workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Verify that the VBA project is signed
                if (vbaProject != null && vbaProject.IsSigned)
                {
                    // Retrieve the raw certificate data
                    byte[] certData = vbaProject.CertRawData;

                    // Ensure certificate data exists before writing
                    if (certData != null && certData.Length > 0)
                    {
                        File.WriteAllBytes(certificateOutputPath, certData);
                        Console.WriteLine($"Certificate exported successfully to '{certificateOutputPath}'.");
                    }
                    else
                    {
                        Console.WriteLine("Certificate data is empty; nothing to export.");
                    }
                }
                else
                {
                    Console.WriteLine("The VBA project is not signed or not present; no certificate to export.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificate.Run();
        }
    }
}
