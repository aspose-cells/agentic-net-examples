using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsExamples
{
    public class DigitalSignatureWithExceptionHandlingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample content
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Document requiring digital signature");

                // Certificate file path and password
                string certPath = "mycertificate.pfx";
                string certPassword = "123456";

                // Ensure the certificate file exists
                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                // Load the certificate
                X509Certificate2 certificate = new X509Certificate2();
                certificate.Import(certPath, certPassword, X509KeyStorageFlags.DefaultKeySet);

                // Create a digital signature instance
                DigitalSignature signature = new DigitalSignature(
                    certificate,
                    "Approved by QA",
                    DateTime.UtcNow);

                // Add the signature to a collection
                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                signatures.Add(signature);

                // Apply the digital signature to the workbook
                workbook.SetDigitalSignature(signatures);

                // Save the signed workbook
                string outputPath = "SignedWorkbook.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine("Workbook signed and saved successfully.");
            }
            catch (Exception ex)
            {
                // Log detailed error information
                Console.WriteLine("An error occurred during the signing process:");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine($"StackTrace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                }
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DigitalSignatureWithExceptionHandlingDemo.Run();
        }
    }
}