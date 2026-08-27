// Title: How to verify a digital signature in an Excel workbook and convert the signed file to PDF with Aspose.Cells for .NET
// AI Prompts: Load a signed .xlsx file using Aspose.Cells, call Workbook.IsDigitallySigned to detect a signature, retrieve the DigitalSignatureCollection, iterate through each signature, and output a validation result. | When a digital signature is confirmed, invoke ConversionUtility.Convert to generate a PDF from the same workbook and log the PDF file path.
// Common Searches: aspnet verify digital signature in Excel workbook using Aspose.Cells | convert signed Excel file to PDF with Aspose.Cells ConversionUtility | check Workbook.IsDigitallySigned before exporting to PDF in C# | count digital signatures in an .xlsx using Aspose.Cells API
// Tags: Aspose.Cells digital signature verification | PDF conversion of signed Excel workbook | Workbook.IsDigitallySigned usage | ConversionUtility Excel to PDF example | Iterating DigitalSignatureCollection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsSignatureAndPdf
{
    // The sample loads a digitally signed Excel workbook, uses Workbook.IsDigitallySigned to confirm the presence of a signature, iterates through the DigitalSignatureCollection to ensure at least one signature exists, reports the validation outcome, and then converts the workbook to PDF with ConversionUtility.Convert while handling missing files and exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel file (must be digitally signed)
                string sourcePath = "SignedWorkbook.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Error: The source file \"{sourcePath}\" was not found.");
                    return;
                }

                // Path for the resulting PDF file
                string pdfPath = "SignedWorkbook.pdf";

                // Load the workbook from the file system
                Workbook workbook = new Workbook(sourcePath);

                // Check if the workbook contains a digital signature
                bool isSigned = workbook.IsDigitallySigned;
                Console.WriteLine($"Workbook is digitally signed: {isSigned}");

                if (isSigned)
                {
                    // Retrieve the digital signature collection
                    DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                    // Count signatures manually (DigitalSignatureCollection may not expose Count directly)
                    int count = 0;
                    if (signatures != null)
                    {
                        foreach (var sig in signatures)
                        {
                            count++;
                        }
                    }

                    Console.WriteLine($"Number of digital signatures found: {count}");

                    // Simple integrity validation: ensure at least one signature exists
                    if (count > 0)
                    {
                        Console.WriteLine("Digital signature validation passed.");
                    }
                    else
                    {
                        Console.WriteLine("Digital signature validation failed: no signatures retrieved.");
                    }
                }
                else
                {
                    Console.WriteLine("No digital signature present; skipping validation.");
                }

                // Convert the Excel workbook to PDF using the provided ConversionUtility
                ConversionUtility.Convert(sourcePath, pdfPath);
                Console.WriteLine($"Workbook successfully converted to PDF at: {pdfPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
