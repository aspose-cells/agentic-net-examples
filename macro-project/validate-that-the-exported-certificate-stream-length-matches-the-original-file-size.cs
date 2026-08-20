// Title: C# – Verify VBA Project Certificate Stream Length Matches Workbook Size Using Aspose.Cells
// Description: Loads a signed .xlsm workbook with Aspose.Cells, reads the original file size, extracts the raw VBA project certificate via workbook.VbaProject.CertRawData, compares the byte‑array length to the workbook size, and reports whether the lengths are identical.
// Keywords: Aspose.Cells VBA certificate | C# workbook size validation | VbaProject CertRawData length | signed .xlsm verification | certificate stream size check
// Common Searches: Aspose.Cells compare certificate length to file size | C# verify signed VBA project size | how to get VBA certificate raw data Aspose | check if workbook certificate matches original size | validate signed Excel macro file integrity
// Developer Intent: Confirm that the exported VBA certificate byte count equals the original workbook file size.
// Use Cases: Detect tampered or corrupted signed workbooks by flagging size mismatches. | Automate compliance logs that record certificate‑to‑file size verification results. | Gate downstream processing to only accept workbooks that pass the certificate size check.
// AI Prompts: Write a C# method using Aspose.Cells that returns true when workbook.VbaProject.CertRawData.Length matches the .xlsm file size. | Generate code that throws a custom exception if the VBA certificate stream length differs from the original workbook size. | Create a logging snippet that records both the workbook size and certificate length and issues a warning on mismatch.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsCertificateValidation
{
    // Loads a signed .xlsm workbook with Aspose.Cells, reads the original file size, extracts the raw VBA project certificate via workbook.VbaProject.CertRawData, compares the byte‑array length to the workbook size, and reports whether the lengths are identical.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains a signed VBA project
            string workbookPath = "signedWorkbook.xlsm";

            // Get the original file size in bytes
            long originalFileSize = new FileInfo(workbookPath).Length;

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Check if the VBA project is signed
            if (workbook.VbaProject.IsSigned)
            {
                // Retrieve the raw certificate data
                byte[] certData = workbook.VbaProject.CertRawData;

                // Determine the length of the exported certificate stream
                long certLength = certData?.Length ?? 0;

                // Output the sizes
                Console.WriteLine($"Original workbook size: {originalFileSize} bytes");
                Console.WriteLine($"Exported certificate stream length: {certLength} bytes");

                // Validate that the lengths match
                if (certLength == originalFileSize)
                {
                    Console.WriteLine("Validation succeeded: lengths match.");
                }
                else
                {
                    Console.WriteLine("Validation failed: lengths do not match.");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a signed VBA project.");
            }
        }
    }
}
