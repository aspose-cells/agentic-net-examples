// Title: Export VBA Project Certificate to a .cer File with Aspose.Cells for .NET (C#)
// Description: Loads an .xlsm workbook, checks for a signed VBA project, extracts the certificate's raw byte array via VbaProject.CertRawData, writes the bytes to a .cer file with File.WriteAllBytes, and saves the workbook. Includes handling for missing files, unsigned projects, and empty certificate data.
// Keywords: Aspose.Cells | C# VBA certificate export | VbaProject CertRawData | extract VBA project certificate | write .cer file | signed VBA project | Aspose.Cells .NET | certificate byte array | Excel macro security | export VBA certificate
// Common Searches: Aspose.Cells export VBA certificate C# | How to get VBA project certificate bytes .NET | Save VBA certificate as .cer file using Aspose | Extract signed macro certificate from Excel workbook | Retrieve CertRawData with Aspose.Cells
// Developer Intent: Extract the signed VBA project's certificate from an Excel workbook and save it as a .cer file.
// Use Cases: Distribute the extracted certificate to external parties for signature verification. | Archive the certificate for compliance auditing or record‑keeping. | Programmatically compare the exported .cer file with a trusted store to validate macro authenticity.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect a signed VBA project in a workbook and export its certificate to a .cer file. | Show how to read CertRawData from a VbaProject, store it in a byte array, and safely write the array to disk. | Explain error‑handling strategies when the VBA project is unsigned, the workbook is missing, or the certificate data is empty while using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads an .xlsm workbook, checks for a signed VBA project, extracts the certificate's raw byte array via VbaProject.CertRawData, writes the bytes to a .cer file with File.WriteAllBytes, and saves the workbook. Includes handling for missing files, unsigned projects, and empty certificate data.
    public class ExportVbaCertificateDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to a workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWorkbook.xlsm";

            // Verify the input file exists to avoid FileNotFoundException
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"Input file not found: '{signedWorkbookPath}'.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(signedWorkbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Verify that the VBA project exists and is signed
                if (vbaProject != null && vbaProject.IsSigned)
                {
                    // Retrieve the raw certificate data as a byte array
                    byte[] certData = vbaProject.CertRawData;

                    // Ensure we have data before writing
                    if (certData != null && certData.Length > 0)
                    {
                        // Export the certificate bytes to an external .cer file
                        string outputCertPath = "VbaCertificate.cer";
                        File.WriteAllBytes(outputCertPath, certData);
                        Console.WriteLine($"Certificate exported to '{outputCertPath}'. Size: {certData.Length} bytes.");
                    }
                    else
                    {
                        Console.WriteLine("Certificate data is empty.");
                    }
                }
                else
                {
                    Console.WriteLine("The workbook does not contain a signed VBA project.");
                }

                // Save the workbook (demonstrates using the Workbook.Save(string) rule)
                string outputWorkbookPath = "ExportVbaCertificateDemo.xlsx";
                workbook.Save(outputWorkbookPath);
                Console.WriteLine($"Workbook saved to '{outputWorkbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
