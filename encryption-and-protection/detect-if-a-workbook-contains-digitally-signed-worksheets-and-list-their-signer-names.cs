// Title: Detect and List Digital Signature Signers in an Excel Workbook using Aspose.Cells for .NET (C#)
// Description: Loads an Excel file with Aspose.Cells, checks the IsDigitallySigned flag, retrieves the DigitalSignatureCollection via GetDigitalSignature, and iterates each DigitalSignature to output the signer’s certificate subject, comments, UTC sign time and validity status.
// Keywords: Aspose.Cells | C# | digital signature detection | list signer names | IsDigitallySigned | GetDigitalSignature | Excel workbook verification | certificate subject extraction | signature validation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells check if workbook is signed | C# get signer name from Excel digital signature | How to read digital signatures in Excel using Aspose | Retrieve certificate subject from Aspose.Cells digital signature | Validate Excel digital signatures .NET
// Developer Intent: Determine whether an Excel workbook contains digital signatures and retrieve each signer’s name along with signature metadata.
// Use Cases: Audit incoming Excel reports to confirm authenticity and record signer information for compliance. | Automate workflows that process only digitally signed spreadsheets, extracting signer details before data transformation. | Generate logs or CSV reports of signature details (signer, comment, timestamp, validity) for regulatory review. | Filter out unsigned workbooks in batch processing pipelines to prevent unauthorized data handling.
// AI Prompts: Create a C# method using Aspose.Cells that returns a list of signer names from all digital signatures in a given workbook. | Write code that validates each digital signature in an Excel file, logs any invalid signatures, and prints signer, comment, and UTC timestamp. | Develop a utility that checks if a workbook is digitally signed and, when true, exports signer names, comments, and sign times to a CSV file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    // Loads an Excel file with Aspose.Cells, checks the IsDigitallySigned flag, retrieves the DigitalSignatureCollection via GetDigitalSignature, and iterates each DigitalSignature to output the signer’s certificate subject, comments, UTC sign time and validity status.
    public class DetectDigitalSignatures
    {
        public static void Run(string workbookPath)
        {
            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(workbookPath);

                // Check if the workbook is digitally signed
                bool isSigned = workbook.IsDigitallySigned;
                Console.WriteLine($"Workbook is digitally signed: {isSigned}");

                if (!isSigned)
                {
                    // No signatures to process
                    return;
                }

                // Retrieve the collection of digital signatures attached to the workbook
                DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                if (signatures == null)
                {
                    Console.WriteLine("No digital signatures collection found in the workbook.");
                    return;
                }

                // Iterate through each signature and display signer information
                int index = 1;
                foreach (DigitalSignature signature in signatures)
                {
                    // The Certificate property holds the X509Certificate2 used for signing
                    var certificate = signature.Certificate;

                    // Extract a readable signer name from the certificate (Subject contains the name)
                    string signerName = certificate != null ? certificate.Subject : "Unknown Signer";

                    Console.WriteLine($"Signature {index}:");
                    Console.WriteLine($"  Signer: {signerName}");
                    Console.WriteLine($"  Comments: {signature.Comments}");
                    Console.WriteLine($"  Sign Time (UTC): {signature.SignTime:u}");
                    Console.WriteLine($"  Is Valid: {signature.IsValid}");
                    index++;
                }

                if (index == 1)
                {
                    Console.WriteLine("No digital signatures found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }

        // Example usage
        public static void Main()
        {
            // Replace with the path to your workbook file
            string path = "SignedWorkbook.xlsx";

            if (!File.Exists(path))
            {
                Console.WriteLine($"File not found: {path}");
                return;
            }

            Run(path);
        }
    }
}
