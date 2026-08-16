// Title: C# Example: Apply a Digital Signature to an Aspose.Cells Workbook with Robust Try‑Catch Error Handling
// Description: The sample creates a new Workbook, inserts sample text, verifies the presence of an X509 certificate file, loads the certificate, builds a DigitalSignature, adds it to a DigitalSignatureCollection, applies the collection via SetDigitalSignature, and saves the signed file. Both the signing operation and the save step are wrapped in try‑catch blocks that output the exception type, message, and stack trace.
// Keywords: Aspose.Cells digital signature C# | try catch exception handling | X509Certificate2 loading | SetDigitalSignature error logging | save signed Excel workbook exception | certificate file not found handling | console logging stack trace
// Common Searches: Aspose.Cells digital signature try catch C# | How to handle errors when signing Excel with Aspose.Cells | C# catch exception SetDigitalSignature | Log stack trace for Aspose.Cells signing failures | Save signed workbook exception handling Aspose.Cells
// Developer Intent: Show how to protect the signing and saving phases of an Aspose.Cells workbook with try‑catch blocks and detailed logging of any runtime errors.
// Use Cases: Validate that the certificate file exists and throw a clear FileNotFoundException before attempting to sign. | Capture and log signing failures caused by an invalid or expired certificate. | Handle I/O exceptions when writing the signed workbook to disk. | Provide concise console feedback for automated build or CI pipelines.
// AI Prompts: Generate C# code that signs an Excel workbook using Aspose.Cells, checks the certificate file, and logs exception type, message, and stack trace. | Refactor the signing logic into a reusable method that returns a boolean success flag and includes comprehensive error handling for both signing and saving. | Create a unit test that verifies the exception handling behavior when the certificate path is incorrect in the Aspose.Cells digital signature example.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsExamples
{
    // The sample creates a new Workbook, inserts sample text, verifies the presence of an X509 certificate file, loads the certificate, builds a DigitalSignature, adds it to a DigitalSignatureCollection, applies the collection via SetDigitalSignature, and saves the signed file. Both the signing operation and the save step are wrapped in try‑catch blocks that output the exception type, message, and stack trace.
    public class DigitalSignatureWithExceptionHandlingDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample content
            workbook.Worksheets[0].Cells["A1"].PutValue("Document requiring digital signature");

            try
            {
                // Path to certificate file
                string certPath = "mycertificate.pfx";
                string certPassword = "password";

                // Verify certificate file exists
                if (!File.Exists(certPath))
                {
                    throw new FileNotFoundException($"Certificate file not found: {certPath}");
                }

                // Load the certificate
                X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);

                // Create digital signature
                DigitalSignature signature = new DigitalSignature(
                    certificate,
                    "Approved by QA Team",
                    DateTime.UtcNow);

                // Add signature to collection
                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                signatures.Add(signature);

                // Apply signature to workbook
                workbook.SetDigitalSignature(signatures);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during digital signing: {ex.GetType().Name} - {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
                return;
            }

            // Save signed workbook
            string outputPath = "SignedWorkbook.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook signed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.GetType().Name} - {ex.Message}");
            }
        }

        // Entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
