// Title: Export VBA Signing Certificate from an Excel Workbook and Load into X509Certificate2 with Aspose.Cells (.NET)
// Description: This example loads a .xlsm file, checks if its VBA project is signed, extracts the raw certificate bytes via VbaProject.CertRawData, creates an X509Certificate2 object directly from the data, and prints key properties such as Subject, Issuer, Thumbprint, and validity period using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | VBA certificate extraction | X509Certificate2 | Excel macro signing | VbaProject.CertRawData | digital signature inspection | Excel workbook security | certificate validation
// Common Searches: How to extract VBA signing certificate from an Excel file using Aspose.Cells | C# load X509Certificate2 from VBA certificate raw data | Check if VBA project is signed with Aspose.Cells .NET | Retrieve VBA project certificate bytes in C# | Inspect Excel macro certificate with Aspose.Cells
// Developer Intent: Obtain the signing certificate of a VBA project embedded in an Excel workbook and examine its details programmatically.
// Use Cases: Confirm that macros are signed by a trusted authority before automated processing. | Log certificate fields (subject, issuer, thumbprint, dates) for compliance audits. | Match the extracted thumbprint against an approved whitelist of certificates. | Diagnose macro‑related security warnings by reviewing certificate information.
// AI Prompts: Write C# code that saves the VBA certificate raw data to a temporary .cer file, then loads it with X509Certificate2 for further analysis. | Provide error‑handling logic for unsigned VBA projects or empty certificate data, including user‑friendly messages. | Explain how to validate the certificate chain and check revocation status after loading the VBA certificate in .NET.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // This example loads a .xlsm file, checks if its VBA project is signed, extracts the raw certificate bytes via VbaProject.CertRawData, creates an X509Certificate2 object directly from the data, and prints key properties such as Subject, Issuer, Thumbprint, and validity period using Aspose.Cells for .NET.
    public class VbaCertificateExportDemo
    {
        public static void Run()
        {
            // Path to an existing workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWithVba.xlsm";

            // Verify that the workbook file exists
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{signedWorkbookPath}'.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(signedWorkbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Check if the VBA project is signed
                if (vbaProject.IsSigned)
                {
                    // Retrieve the raw certificate data
                    byte[] certData = vbaProject.CertRawData;

                    if (certData != null && certData.Length > 0)
                    {
                        // Load the certificate directly from the raw data (no temporary file needed)
                        X509Certificate2 certificate = new X509Certificate2(certData);

                        // Display some certificate information
                        Console.WriteLine($"Subject: {certificate.Subject}");
                        Console.WriteLine($"Issuer: {certificate.Issuer}");
                        Console.WriteLine($"Thumbprint: {certificate.Thumbprint}");
                        Console.WriteLine($"Valid From: {certificate.NotBefore}");
                        Console.WriteLine($"Valid To: {certificate.NotAfter}");
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
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            VbaCertificateExportDemo.Run();
        }
    }
}
