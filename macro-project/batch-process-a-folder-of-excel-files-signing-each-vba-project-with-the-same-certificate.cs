using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaBatchSigner
{
    public class VbaBatchSigner
    {
        /// <summary>
        /// Signs all VBA projects in macro‑enabled Excel files within a folder using the same certificate.
        /// </summary>
        /// <param name="sourceFolder">Folder containing the source .xlsm files.</param>
        /// <param name="outputFolder">Folder where signed files will be saved.</param>
        /// <param name="certificatePath">Path to the .pfx certificate file.</param>
        /// <param name="certificatePassword">Password for the certificate.</param>
        public static void SignVbaProjectsInFolder(string sourceFolder, string outputFolder, string certificatePath, string certificatePassword)
        {
            // Verify certificate file exists
            if (!File.Exists(certificatePath))
            {
                Console.WriteLine($"Certificate file not found: {certificatePath}");
                return;
            }

            X509Certificate2 certificate;
            try
            {
                // Load the signing certificate once
                certificate = new X509Certificate2(certificatePath, certificatePassword);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load certificate: {ex.Message}");
                return;
            }

            // Create a DigitalSignature instance that will be reused for each workbook
            DigitalSignature digitalSignature = new DigitalSignature(certificate, "Batch VBA Signature", DateTime.Now);

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Process each .xlsm file in the source folder
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsm"))
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Source file not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (macro‑enabled)
                    Workbook workbook = new Workbook(filePath);

                    // Access the VBA project
                    VbaProject vbaProject = workbook.VbaProject;

                    // If a VBA project exists, sign it
                    if (vbaProject != null)
                    {
                        vbaProject.Sign(digitalSignature);
                    }

                    // Determine output file path (preserve original file name)
                    string outputFilePath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                    // Save the signed workbook as macro‑enabled format
                    workbook.Save(outputFilePath, SaveFormat.Xlsm);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }

        // Example usage
        public static void Main()
        {
            try
            {
                string sourceFolder = @"C:\ExcelFiles\Source";
                string outputFolder = @"C:\ExcelFiles\Signed";
                string certPath = @"C:\Certificates\mycert.pfx";
                string certPassword = "certPassword";

                SignVbaProjectsInFolder(sourceFolder, outputFolder, certPath, certPassword);

                Console.WriteLine("Batch signing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}