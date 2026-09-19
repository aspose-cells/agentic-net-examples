// Title: How to export a VBA project's digital certificate to a binary file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, uses reflection to read the VBA project's Signature property, extracts the X509Certificate, and writes the DER‑encoded bytes to a .bin file with comprehensive error handling. | Create a method that verifies the existence of a VBA project and its digital signature in a workbook, obtains the certificate object, and saves the certificate data to a user‑specified output path. | Generate a console application example that backs up a VBA macro's digital certificate from an Excel file using Aspose.Cells, including logging for missing project, missing signature, and file‑write failures.
// Common Searches: Aspose.Cells C# export VBA macro certificate to file | retrieve X509Certificate from Excel VBA project using .NET | backup VBA project digital signature with Aspose.Cells | save VBA project's digital certificate as binary using C# reflection
// Tags: export VBA certificate Aspose.Cells C# | retrieve VbaProject signature X509Certificate | save certificate to binary file .NET | reflection access VbaProject.Signature | backup VBA macro digital signature

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The program loads an Excel workbook, accesses its VBA project via Aspose.Cells, uses reflection to obtain the VBA signature and associated X509Certificate, then exports the certificate as a DER‑encoded binary file while handling cases where the project, signature, or certificate is missing.
class VbaCertificateExporter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "VbaCertificate.bin";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains the VBA project
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            var vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Attempt to retrieve the VBA signature via reflection (avoids compile‑time dependency on VbaSignature)
            object vbaSignature = null;
            try
            {
                var signatureProp = vbaProject.GetType().GetProperty("Signature");
                vbaSignature = signatureProp?.GetValue(vbaProject);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to access VBA signature: {ex.Message}");
            }

            if (vbaSignature == null)
            {
                Console.WriteLine("No VBA digital signature found in the workbook.");
                return;
            }

            // Attempt to get the certificate from the signature object
            X509Certificate certificate = null;
            try
            {
                var certProp = vbaSignature.GetType().GetProperty("Certificate");
                certificate = certProp?.GetValue(vbaSignature) as X509Certificate;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving certificate: {ex.Message}");
            }

            if (certificate == null)
            {
                Console.WriteLine("No VBA digital certificate found in the workbook.");
                return;
            }

            // Export the certificate to a byte array (DER encoded)
            byte[] certificateBytes = certificate.Export(X509ContentType.Cert);

            // Write the certificate bytes to the output file
            try
            {
                File.WriteAllBytes(outputPath, certificateBytes);
                Console.WriteLine("VBA digital certificate exported successfully.");
            }
            catch (Exception writeEx)
            {
                Console.WriteLine($"Failed to write the certificate file: {writeEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
