using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaCertificateExportDemo
    {
        public static void Main()
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
            // Path to an existing workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWithVba.xlsm";

            // Verify the workbook file exists
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"Workbook file not found: {signedWorkbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(signedWorkbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Ensure the VBA project is signed
            if (!vbaProject.IsSigned)
            {
                Console.WriteLine("The VBA project is not signed. No certificate to export.");
                return;
            }

            // Retrieve the raw certificate data
            byte[] certData = vbaProject.CertRawData;

            if (certData == null || certData.Length == 0)
            {
                Console.WriteLine("Certificate raw data is empty.");
                return;
            }

            // Load the certificate directly from the raw data (no temporary file needed)
            X509Certificate2 certificate;
            try
            {
                certificate = new X509Certificate2(certData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load certificate from raw data: {ex.Message}");
                return;
            }

            // Display some certificate information
            Console.WriteLine($"Certificate Subject: {certificate.Subject}");
            Console.WriteLine($"Certificate Issuer: {certificate.Issuer}");
            Console.WriteLine($"Certificate Thumbprint: {certificate.Thumbprint}");
            Console.WriteLine($"Valid From: {certificate.NotBefore}");
            Console.WriteLine($"Valid To: {certificate.NotAfter}");
        }
    }
}