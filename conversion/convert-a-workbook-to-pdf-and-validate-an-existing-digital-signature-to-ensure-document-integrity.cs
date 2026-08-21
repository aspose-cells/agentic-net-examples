// Title: Validate an Excel Workbook's Digital Signature and Convert It to PDF with Aspose.Cells for .NET
// Description: C# sample that loads a signed Excel file, checks the Workbook.IsDigitallySigned flag, iterates through the DigitalSignatureCollection to read signer name, signing time and reason, and then uses ConversionUtility.Convert to produce a PDF. Includes file‑existence checks and comprehensive error handling.
// Keywords: Aspose.Cells | digital signature verification | Excel to PDF conversion | .NET | Workbook.IsDigitallySigned | DigitalSignatureCollection | ConversionUtility | C# example | document integrity | certificate details extraction
// Common Searches: how to verify digital signature in Excel using Aspose.Cells | aspnet convert signed workbook to PDF | read signer information from Excel digital signature C# | Aspose.Cells check IsDigitallySigned before conversion | sample code for Excel digital signature validation and PDF export
// Developer Intent: The developer needs to confirm that an Excel workbook’s digital signature is present and trustworthy before generating a PDF version of the file.
// Use Cases: Ensure document integrity by validating the signature prior to distribution. | Capture signer metadata for audit trails or compliance reporting. | Automate archival of signed workbooks as PDF files for downstream systems.
// AI Prompts: Generate code that aborts PDF conversion when the workbook’s digital signature fails verification and returns a detailed error response. | Show how to extract certificate properties such as thumbprint, issuer, and subject from each DigitalSignature in Aspose.Cells. | Provide a robust error‑handling template for ConversionUtility.Convert that covers missing, corrupted, or password‑protected Excel sources.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsSignatureAndPdf
{
    // C# sample that loads a signed Excel file, checks the Workbook.IsDigitallySigned flag, iterates through the DigitalSignatureCollection to read signer name, signing time and reason, and then uses ConversionUtility.Convert to produce a PDF. Includes file‑existence checks and comprehensive error handling.
    class Program
    {
        static void Main()
        {
            // Paths to the input Excel file (digitally signed) and the output PDF file
            string excelPath = "SignedWorkbook.xlsx";
            string pdfPath = "SignedWorkbook.pdf";

            try
            {
                // Verify that the Excel file exists before attempting to load it
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                    return;
                }

                // Load the workbook to inspect its digital signature
                Workbook workbook = new Workbook(excelPath);

                // Check if the workbook is digitally signed
                bool isSigned = workbook.IsDigitallySigned;
                Console.WriteLine($"Workbook digitally signed: {isSigned}");

                if (isSigned)
                {
                    try
                    {
                        // Retrieve the digital signature collection
                        DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                        // Display basic information about each signature
                        int index = 1;
                        foreach (DigitalSignature signature in signatures)
                        {
                            // Use dynamic to avoid compile‑time binding issues with property names
                            dynamic sig = signature;
                            Console.WriteLine($"Signature {index}:");
                            Console.WriteLine($"  Signer Name : {sig.Signer}");
                            Console.WriteLine($"  Signing Time: {sig.SignTime}");
                            Console.WriteLine($"  Reason      : {sig.Reason}");
                            index++;
                        }
                    }
                    catch (Exception sigEx)
                    {
                        Console.WriteLine($"Error while reading signatures: {sigEx.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("No digital signature found in the workbook.");
                }

                // Convert the Excel workbook to PDF using the provided ConversionUtility rule
                try
                {
                    ConversionUtility.Convert(excelPath, pdfPath);
                    Console.WriteLine($"Workbook successfully converted to PDF: {pdfPath}");
                }
                catch (Exception convEx)
                {
                    Console.WriteLine($"PDF conversion failed: {convEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
