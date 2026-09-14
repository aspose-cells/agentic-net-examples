// Title: Extract and save the X509 certificate from a digitally signed Excel workbook to a .cer file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens a signed .xlsx workbook with Aspose.Cells, uses reflection to access its DigitalSignatureCollection, and obtains the X509Certificate2 of the first signature. | Show how to export the obtained X509 certificate to a DER‑encoded byte array and persist it as a .cer file on disk. | Create a reusable method that accepts an input workbook path and an output certificate path, extracts the certificate, handles missing signatures, and writes the byte array to a file.
// Common Searches: how to get X509 certificate from a signed Excel file using Aspose.Cells C# | Aspose.Cells export digital signature certificate to .cer file | C# extract first digital signature certificate from workbook with reflection | save certificate bytes from Excel digital signature using Aspose.Cells .NET
// Tags: Aspose.Cells digital signature extraction C# | export X509 certificate to .cer file .NET | reflection access DigitalSignatureCollection Aspose.Cells | write certificate byte array to disk C# | retrieve certificate from signed workbook Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures; // Required for DigitalSignature classes

// The example loads a signed Excel workbook with Aspose.Cells, uses reflection to obtain its DigitalSignatureCollection, extracts the X509Certificate2 from the first DigitalSignature, exports the certificate as a DER‑encoded byte array, and writes the bytes to a .cer file for external distribution.
class ExportCertificate
{
    static void Main()
    {
        try
        {
            // Path to the signed workbook
            string inputPath = "SignedWorkbook.xlsx";

            // Verify the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains a digital signature
            Workbook workbook = new Workbook(inputPath);

            // Use reflection to obtain the DigitalSignatureCollection (avoids compile‑time dependency on specific API versions)
            PropertyInfo sigCollectionProp = workbook.GetType().GetProperty("DigitalSignatureCollection");
            if (sigCollectionProp == null)
            {
                Console.WriteLine("Digital signatures are not supported by the current Aspose.Cells version.");
                return;
            }

            object signaturesObj = sigCollectionProp.GetValue(workbook);
            if (signaturesObj == null)
            {
                Console.WriteLine("No digital signatures found in the workbook.");
                return;
            }

            // Get the Count property of the collection
            PropertyInfo countProp = signaturesObj.GetType().GetProperty("Count");
            int count = (int)(countProp?.GetValue(signaturesObj) ?? 0);
            if (count == 0)
            {
                Console.WriteLine("No digital signatures found in the workbook.");
                return;
            }

            // Retrieve the first DigitalSignature (using the indexer property)
            PropertyInfo indexerProp = signaturesObj.GetType().GetProperty("Item");
            object signatureObj = indexerProp?.GetValue(signaturesObj, new object[] { 0 });
            if (signatureObj == null)
            {
                Console.WriteLine("Unable to retrieve the digital signature.");
                return;
            }

            DigitalSignature signature = signatureObj as DigitalSignature;
            if (signature == null)
            {
                Console.WriteLine("The retrieved object is not a DigitalSignature.");
                return;
            }

            // Retrieve the X509 certificate associated with the signature
            X509Certificate2 certificate = signature.Certificate;
            if (certificate == null)
            {
                Console.WriteLine("The signature does not contain an X509 certificate.");
                return;
            }

            // Export the certificate to a byte array (DER encoded)
            byte[] certificateBytes = certificate.Export(X509ContentType.Cert);

            // Write the byte array to a file for external distribution
            string outputPath = "ExportedCertificate.cer";
            File.WriteAllBytes(outputPath, certificateBytes);

            Console.WriteLine($"Certificate exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
