// Title: C# Batch Sign VBA Projects in Multiple .xlsm Workbooks with Aspose.Cells
// Description: A C# console utility that scans a folder for macro‑enabled Excel files (*.xlsm), loads each workbook with Aspose.Cells, detects a VBA project, and digitally signs it using a single X509 certificate from a PFX file. Signed files are saved to a target directory while preserving original names, enabling automated compliance, template preparation, and CI/CD signing.
// Keywords: Aspose.Cells VBA signing | C# batch digital signature Excel | sign .xlsm files programmatically | X509 certificate Excel macros | macro-enabled workbook signing | automated VBA project signing | digital signature Aspose.Cells .NET
// Common Searches: C# code to sign VBA project in Excel | batch sign .xlsm files Aspose.Cells | apply same certificate to multiple Excel workbooks | automate VBA digital signature .NET | sign macro-enabled Excel files from folder
// Developer Intent: Automatically apply a single digital certificate to the VBA projects of all macro‑enabled Excel workbooks in a specified folder.
// Use Cases: Ensure regulatory compliance for financial reports that contain macros. | Distribute trusted Excel templates across an organization. | Add a signing step to build pipelines that generate .xlsm reports. | Protect macro integrity before sending files to external partners.
// AI Prompts: Generate C# code that signs VBA projects in .xlsm files using a PFX certificate and Aspose.Cells, handling missing VBA projects gracefully. | Refactor the program to accept source folder, output folder, certificate path, and password via command‑line arguments and produce a JSON summary of processed files. | Create NUnit unit tests for the batch signing logic, mocking the X509Certificate2 and verifying that workbooks without VBA are left unchanged.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace BatchVbaSigning
{
    // A C# console utility that scans a folder for macro‑enabled Excel files (*.xlsm), loads each workbook with Aspose.Cells, detects a VBA project, and digitally signs it using a single X509 certificate from a PFX file. Signed files are saved to a target directory while preserving original names, enabling automated compliance, template preparation, and CI/CD signing.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the Excel files to be signed
            string sourceFolder = @"C:\InputFolder";

            // Folder where signed files will be saved
            string outputFolder = @"C:\SignedFolder";

            // Path to the certificate (PFX) and its password
            string certificatePath = @"C:\Certificates\mycert.pfx";
            string certificatePassword = "yourPassword";

            try
            {
                // Ensure the output directory exists
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // Verify certificate file exists
                if (!File.Exists(certificatePath))
                {
                    Console.WriteLine($"Certificate file not found: {certificatePath}");
                    return;
                }

                // Load the certificate once for all files
                X509Certificate2 certificate = new X509Certificate2(certificatePath, certificatePassword, X509KeyStorageFlags.MachineKeySet);

                // Create a DigitalSignature instance that will be applied to each VBA project
                DigitalSignature digitalSignature = new DigitalSignature(certificate, "Batch VBA signing", DateTime.Now);

                // Process each macro-enabled workbook in the source folder
                foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsm"))
                {
                    try
                    {
                        // Verify workbook file exists (redundant but safe)
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"Workbook file not found: {filePath}");
                            continue;
                        }

                        // Load the workbook from file
                        Workbook workbook = new Workbook(filePath);

                        // Access the VBA project; it may be null if the workbook has no VBA
                        VbaProject vbaProject = workbook.VbaProject;

                        if (vbaProject != null)
                        {
                            // Sign the VBA project with the prepared digital signature
                            vbaProject.Sign(digitalSignature);
                        }

                        // Determine the output file path
                        string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                        // Save the signed workbook as a macro-enabled file
                        workbook.Save(outputPath, SaveFormat.Xlsm);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch signing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
