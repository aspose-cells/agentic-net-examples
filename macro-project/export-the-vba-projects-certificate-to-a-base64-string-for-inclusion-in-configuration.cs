// Title: Export a Signed VBA Project Certificate to Base64 Using Aspose.Cells for .NET
// Description: Loads an Excel workbook, accesses its VBA project, checks the IsSigned flag, extracts the raw certificate bytes (CertRawData), converts them to a Base64 string, prints the result, and writes it to a UTF‑8 text file for configuration or audit purposes.
// Keywords: Aspose.Cells VBA certificate export | C# extract VBA project certificate | Convert VBA certificate to Base64 | CertRawData Aspose.Cells | Save VBA certificate to config file | Signed macro certificate extraction | Excel VBA project signing
// Common Searches: how to get VBA project certificate with Aspose.Cells | export signed VBA macro certificate as Base64 C# | Aspose.Cells check if VBA project is signed | retrieve CertRawData from Excel workbook | save VBA certificate to text file .NET
// Developer Intent: Extract the certificate from a signed VBA project and encode it as a Base64 string for storage or further processing.
// Use Cases: Persist the Base64 certificate in application settings to verify macro signatures at runtime. | Include the certificate string in deployment scripts that re‑sign workbooks on different machines. | Log the exported certificate for compliance and audit trails before distributing the workbook.
// AI Prompts: Generate a C# method that loads an .xlsm file with Aspose.Cells, verifies the VBA project is signed, and returns the certificate as a Base64 string. | Create error‑handling code for exporting a VBA project's certificate to a text file, covering missing workbook, unsigned project, and empty certificate scenarios. | Write a snippet that saves the Base64‑encoded VBA certificate into an appsettings.json entry using Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    // Loads an Excel workbook, accesses its VBA project, checks the IsSigned flag, extracts the raw certificate bytes (CertRawData), converts them to a Base64 string, prints the result, and writes it to a UTF‑8 text file for configuration or audit purposes.
    public class CertificateExporter
    {
        // Entry point for the application.
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        // Exports the VBA project's certificate (if signed) to a Base64 string.
        public static void Run()
        {
            // Path to the workbook that contains a signed VBA project.
            string workbookPath = "SignedWorkbook.xlsm";

            // Verify that the workbook file exists to avoid FileNotFoundException.
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project.
                VbaProject vbaProject = workbook.VbaProject;

                // Check if the VBA project is signed.
                if (vbaProject.IsSigned)
                {
                    // Retrieve the raw certificate data.
                    byte[] certData = vbaProject.CertRawData;

                    // Ensure certificate data exists.
                    if (certData != null && certData.Length > 0)
                    {
                        // Convert the raw bytes to a Base64 string.
                        string base64Cert = Convert.ToBase64String(certData);

                        // Output the Base64 string.
                        Console.WriteLine("VBA Project Certificate (Base64):");
                        Console.WriteLine(base64Cert);

                        // Save the Base64 string to a configuration file.
                        string configPath = "VbaCertificateConfig.txt";
                        File.WriteAllText(configPath, base64Cert, Encoding.UTF8);
                        Console.WriteLine($"Base64 certificate saved to: {configPath}");
                    }
                    else
                    {
                        Console.WriteLine("Certificate raw data is empty.");
                    }
                }
                else
                {
                    Console.WriteLine("The VBA project is not signed; no certificate to export.");
                }

                // Optionally, save the workbook if any modifications were made.
                // workbook.Save("ModifiedWorkbook.xlsm", SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
