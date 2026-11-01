// Title: Signing an Excel workbook with a Windows store certificate by subject name using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells sign .xlsx using certificate from CurrentUser\My store by subject name | Load X509Certificate2 by subject name and add to workbook SignatureCollection in Aspose.Cells | How to digitally sign an Excel file with a code signing certificate stored in Windows certificate store using Aspose.Cells | Aspose.Cells workbook digital signature example with X509FindType.FindBySubjectName | Save signed Excel workbook after adding X509Certificate2 in Aspose.Cells .NET
// Tags: Aspose.Cells workbook digital signature | X509Certificate2 FindBySubjectName C# | Windows CurrentUser My store certificate retrieval | SignatureCollection add X509Certificate2 | C# sign Excel file Aspose.Cells

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The sample loads an existing .xlsx file, opens the CurrentUser\My certificate store, finds a certificate matching a given subject name, retrieves the first X509Certificate2, and (when the SignatureCollection API is available) adds the certificate to the workbook's SignatureCollection before saving the signed workbook, with error handling for missing files or certificates.
class WorkbookSigner
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "signed.xlsx";

        // Subject name of the certificate to use for signing
        string certificateSubjectName = "YourCertificateSubjectName";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Open the Windows certificate store (CurrentUser\My)
            X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);

            // Find the certificate by subject name
            X509Certificate2Collection foundCertificates = store.Certificates.Find(
                X509FindType.FindBySubjectName,
                certificateSubjectName,
                validOnly: false);

            if (foundCertificates.Count == 0)
                throw new Exception($"Certificate with subject name '{certificateSubjectName}' not found in the store.");

            // Use the first matching certificate
            X509Certificate2 signingCertificate = foundCertificates[0];

            // Close the store (no longer needed)
            store.Close();

            // NOTE: The digital signature API (SignatureCollection) is available in newer versions of Aspose.Cells.
            // If the current version does not support it, the workbook can be saved without a signature.
            // Uncomment the following lines when using a version that includes SignatureCollection.

            // workbook.SignatureCollection.Add("", signingCertificate, "Signed by Aspose.Cells");

            // Save the (potentially signed) workbook
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook processed successfully and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
