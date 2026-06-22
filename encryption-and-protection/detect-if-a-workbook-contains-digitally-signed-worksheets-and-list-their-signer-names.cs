using System;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using System.Security.Cryptography.X509Certificates;

class DetectDigitalSignatures
{
    static void Main(string[] args)
    {
        // Path to the workbook to be examined
        string workbookPath = "SignedWorkbook.xlsx";

        // Load the workbook (uses the provided Workbook(string) constructor)
        Workbook workbook = new Workbook(workbookPath);

        // Check if the workbook is digitally signed
        bool isSigned = workbook.IsDigitallySigned;
        Console.WriteLine($"Workbook is digitally signed: {isSigned}");

        if (isSigned)
        {
            // Retrieve the collection of digital signatures
            DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

            if (signatures != null && signatures.GetEnumerator().MoveNext())
            {
                Console.WriteLine("Signers found in the workbook:");
                foreach (DigitalSignature signature in signatures)
                {
                    // Each signature contains an X509Certificate2 object
                    X509Certificate2 cert = signature.Certificate;

                    // Extract signer information (Subject contains the signer name)
                    string signerName = cert != null ? cert.Subject : "Unknown signer";

                    Console.WriteLine($"- {signerName}");
                }
            }
            else
            {
                Console.WriteLine("No digital signatures were found despite IsDigitallySigned being true.");
            }
        }

        // Clean up
        workbook.Dispose();
    }
}