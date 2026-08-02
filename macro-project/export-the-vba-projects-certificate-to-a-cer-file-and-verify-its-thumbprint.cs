// Title: Export VBA Project Certificate to a .cer File and Verify Its Thumbprint with Aspose.Cells for .NET
// Description: Loads a workbook that contains a signed VBA project, checks the IsSigned flag, extracts the raw certificate via VbaProject.CertRawData, writes it to a .cer file, loads the certificate with X509Certificate2, displays the thumbprint, recomputes the thumbprint from the raw data, and confirms the values match. The workbook is then saved unchanged.
// Keywords: Aspose.Cells VBA certificate export | VbaProject CertRawData C# | export VBA signing certificate .cer | thumbprint verification .NET | X509Certificate2 VBA macro signature | C# Aspose.Cells signed macro | Windows .NET certificate extraction
// Common Searches: export signed VBA project certificate Aspose.Cells | how to get VBA macro certificate .cer C# | verify VBA project thumbprint with Aspose.Cells | retrieve raw certificate data from VBA project .NET | compare exported certificate thumbprint with raw data
// Developer Intent: Extract a signed VBA project's certificate, save it as a .cer file, and confirm the thumbprint matches the raw data using Aspose.Cells.
// Use Cases: Backup the signing certificate of a VBA macro for archival or distribution. | Validate macro authenticity before automated processing by checking the certificate thumbprint. | Perform compliance audits across multiple workbooks by exporting VBA certificates and logging their thumbprints.
// AI Prompts: Generate C# code that uses Aspose.Cells to extract a VBA project's certificate and save it as a .cer file. | Create a method that takes a workbook path, exports the VBA certificate, and returns true if the exported thumbprint matches the raw-data thumbprint. | Explain how to handle unsigned VBA projects or empty certificate data when using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Security.Cryptography.X509Certificates;

namespace AsposeCellsExamples
{
    // Loads a workbook that contains a signed VBA project, checks the IsSigned flag, extracts the raw certificate via VbaProject.CertRawData, writes it to a .cer file, loads the certificate with X509Certificate2, displays the thumbprint, recomputes the thumbprint from the raw data, and confirms the values match. The workbook is then saved unchanged.
    public class ExportVbaCertificateDemo
    {
        public static void Run()
        {
            try
            {
                // Path to the workbook that contains a signed VBA project
                string signedWorkbookPath = "SignedWithVba.xlsm";

                // Verify the workbook file exists
                if (!File.Exists(signedWorkbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {signedWorkbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(signedWorkbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Check if the VBA project is signed
                if (vbaProject != null && vbaProject.IsSigned)
                {
                    // Retrieve the raw certificate data
                    byte[] certData = vbaProject.CertRawData;

                    if (certData != null && certData.Length > 0)
                    {
                        // Export the certificate to a .cer file
                        string certFilePath = "VbaCertificate.cer";
                        File.WriteAllBytes(certFilePath, certData);
                        Console.WriteLine($"Certificate exported to: {certFilePath}");

                        // Load the exported certificate
                        X509Certificate2 certificate = new X509Certificate2(certFilePath);
                        Console.WriteLine($"Certificate Thumbprint: {certificate.Thumbprint}");

                        // Verify thumbprint by recomputing from the raw data (optional)
                        X509Certificate2 certFromRaw = new X509Certificate2(certData);
                        Console.WriteLine($"Thumbprint from raw data: {certFromRaw.Thumbprint}");

                        // Compare the two thumbprints
                        bool thumbprintsMatch = string.Equals(
                            certificate.Thumbprint,
                            certFromRaw.Thumbprint,
                            StringComparison.OrdinalIgnoreCase);
                        Console.WriteLine($"Thumbprints match: {thumbprintsMatch}");
                    }
                    else
                    {
                        Console.WriteLine("Certificate raw data is empty.");
                    }
                }
                else
                {
                    Console.WriteLine("VBA project is not signed; no certificate to export.");
                }

                // Optionally, save the workbook (unchanged) to demonstrate lifecycle usage
                string outputPath = "ExportVbaCertificateDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
