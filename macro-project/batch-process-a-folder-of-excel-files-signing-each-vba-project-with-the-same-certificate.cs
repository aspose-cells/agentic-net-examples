// Title: Batch sign VBA projects in multiple .xlsm files with Aspose.Cells for .NET (C#)
// Description: A C# console utility that scans a folder for macro‑enabled Excel workbooks, loads each file with Aspose.Cells, signs its VBA project using a single PFX certificate, and writes the signed copies to an output directory. Includes folder validation, certificate loading, and per‑file error handling.
// Keywords: Aspose.Cells VBA signing | batch digital signature Excel | C# sign .xlsm macros | load X509Certificate2 PFX | macro-enabled workbook automation | .NET Excel digital certificate | bulk VBA project signing
// Common Searches: how to batch sign VBA projects in .xlsm files using C# | Aspose.Cells example for signing multiple macro workbooks | C# program to apply the same PFX certificate to many Excel files | automate digital signing of Excel macros with Aspose.Cells | load X509Certificate2 and sign VBA project in a folder
// Developer Intent: Automatically apply a single digital certificate to the VBA projects of all macro‑enabled Excel files in a specified directory.
// Use Cases: Prepare a distribution package of spreadsheets where each macro is trusted by the corporate certificate. | Integrate the tool into a CI/CD pipeline to sign generated .xlsm reports before release. | Re‑sign existing workbooks after a certificate renewal across shared network drives.
// AI Prompts: Write C# code that uses Aspose.Cells to sign the VBA project of every .xlsm file in a folder with a given PFX certificate and password. | Show how to extend the batch signing script to create a CSV log containing file name, signing result, and error details. | Demonstrate how to verify a workbook’s VBA project signature and retrieve its metadata after saving with Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaBatchSigning
{
    // A C# console utility that scans a folder for macro‑enabled Excel workbooks, loads each file with Aspose.Cells, signs its VBA project using a single PFX certificate, and writes the signed copies to an output directory. Includes folder validation, certificate loading, and per‑file error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Input parameters
            string sourceFolder = @"C:\InputExcelFiles";          // Folder containing Excel files to sign
            string outputFolder = @"C:\SignedExcelFiles";         // Folder where signed files will be saved
            string certificatePath = @"C:\Certificates\mycert.pfx"; // Path to the signing certificate (PFX)
            string certificatePassword = "certPassword";          // Password for the PFX file

            // Validate source folder
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Validate certificate file
            if (!File.Exists(certificatePath))
            {
                Console.WriteLine($"Certificate file not found: {certificatePath}");
                return;
            }

            X509Certificate2 signingCertificate;
            try
            {
                // Load the signing certificate
                signingCertificate = new X509Certificate2(certificatePath, certificatePassword, X509KeyStorageFlags.MachineKeySet);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load certificate: {ex.Message}");
                return;
            }

            // Create a DigitalSignature instance that will be reused for all workbooks
            DigitalSignature digitalSignature = new DigitalSignature(signingCertificate, "Batch VBA Signing", DateTime.Now);

            // Process each .xlsm file in the source folder
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsm"))
            {
                try
                {
                    // Load the workbook (macro-enabled)
                    Workbook workbook = new Workbook(filePath);

                    // Access the VBA project
                    VbaProject vbaProject = workbook.VbaProject;

                    // If the workbook contains a VBA project, sign it
                    if (vbaProject != null)
                    {
                        vbaProject.Sign(digitalSignature);
                    }

                    // Determine output file path (same file name in output folder)
                    string outputFilePath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                    // Save the signed workbook as macro-enabled file
                    workbook.Save(outputFilePath, SaveFormat.Xlsm);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch signing completed.");
        }
    }
}
