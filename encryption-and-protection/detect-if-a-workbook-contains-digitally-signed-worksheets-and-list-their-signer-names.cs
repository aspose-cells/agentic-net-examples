// Title: Detect Excel workbook digital signatures and list signer names with Aspose.Cells for .NET
// Description: This C# example loads an Excel file using Aspose.Cells, checks the Workbook.IsDigitallySigned flag, retrieves the DigitalSignatureCollection via GetDigitalSignature, and prints each signer's name extracted from the certificate's Subject field.
// Keywords: Aspose.Cells digital signature detection | Workbook.IsDigitallySigned C# | GetDigitalSignature example | extract signer name from Excel certificate | C# verify Excel workbook signature | .NET Excel security
// Common Searches: how to check if an Excel file is digitally signed using Aspose.Cells | list signers of Excel digital signatures in C# | Aspose.Cells GetDigitalSignature usage | retrieve certificate subject from Excel workbook signature | C# code to read digital signatures in .xlsx
// Developer Intent: Determine whether an Excel workbook is digitally signed and obtain the names of all signers.
// Use Cases: Validate authenticity of incoming Excel reports before automated processing. | Log signer information for compliance and audit trails. | Trigger workflow steps only when a workbook is signed by a trusted certificate.
// AI Prompts: Generate C# code with Aspose.Cells that returns a list of signer names from an Excel file's digital signatures. | Explain how to handle a DigitalSignature object that lacks an associated X509Certificate when extracting signer details. | Create a reusable method that checks Workbook.IsDigitallySigned and returns signer names as a string array.

using System;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using System.Security.Cryptography.X509Certificates;

// This C# example loads an Excel file using Aspose.Cells, checks the Workbook.IsDigitallySigned flag, retrieves the DigitalSignatureCollection via GetDigitalSignature, and prints each signer's name extracted from the certificate's Subject field.
class Program
{
    static void Main()
    {
        // Path to the workbook to be examined
        string workbookPath = "SignedWorkbook.xlsx";

        // Load the workbook (uses the provided Workbook(string) constructor)
        Workbook workbook = new Workbook(workbookPath);

        // Check if the workbook is digitally signed
        Console.WriteLine($"Workbook is digitally signed: {workbook.IsDigitallySigned}");

        if (workbook.IsDigitallySigned)
        {
            // Retrieve the collection of digital signatures
            DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

            int signatureIndex = 1;
            foreach (DigitalSignature signature in signatures)
            {
                // Extract signer information from the certificate (if available)
                string signerName = signature.Certificate?.Subject ?? "Unknown signer";

                Console.WriteLine($"Signature {signatureIndex}: Signer = {signerName}");
                signatureIndex++;
            }
        }
    }
}
