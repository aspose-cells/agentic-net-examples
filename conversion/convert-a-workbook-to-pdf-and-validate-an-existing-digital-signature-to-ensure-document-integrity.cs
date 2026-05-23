using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using Aspose.Cells.Utility;

class WorkbookToPdfWithSignatureValidation
{
    static void Main()
    {
        // Paths for source Excel and output PDF
        string sourceFile = "SignedWorkbook.xlsx";
        string pdfFile = "SignedWorkbook.pdf";

        // Verify source file exists
        if (!File.Exists(sourceFile))
        {
            Console.WriteLine($"Source file not found: {sourceFile}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(sourceFile);

            // Check for digital signatures
            if (workbook.IsDigitallySigned)
            {
                Console.WriteLine("Workbook is digitally signed.");

                // Get the signature collection
                DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                if (signatures != null)
                {
                    int sigCount = 0;
                    foreach (var sigObj in signatures)
                    {
                        sigCount++;
                        dynamic sig = sigObj; // Use dynamic to access members without compile‑time binding
                        Console.WriteLine($"Signature {sigCount}:");
                        Console.WriteLine($"  Signer       : {sig.Signer}");
                        Console.WriteLine($"  Signing Time : {sig.SigningTime}");
                        Console.WriteLine($"  Comment      : {sig.Comment}");
                    }

                    Console.WriteLine($"Number of signatures found: {sigCount}");
                }
                else
                {
                    Console.WriteLine("Signature collection is null.");
                }
            }
            else
            {
                Console.WriteLine("Workbook is NOT digitally signed.");
            }

            // Convert workbook to PDF
            ConversionUtility.Convert(sourceFile, pdfFile);
            Console.WriteLine($"Workbook successfully converted to PDF: {pdfFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}