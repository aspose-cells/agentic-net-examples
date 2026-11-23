// Title: Export a signed VBA project's digital certificate to a PEM file using Aspose.Cells for .NET
// Common Searches: aspocells export vba project certificate to pem c# | retrieve signed vba macro certificate using aspose.cells .net | backup vba digital certificate as pem file in c# | extract certificate from signed vba project with aspose.cells | convert x509certificate2 to pem using c# aspocells
// Tags: Aspose.Cells VBA certificate PEM export workflow | Reflection access to VbaProject SignatureInfo | DER to PEM conversion for X509Certificate2 in C# | Backing up signed VBA macro certificates | Workbook loading and VBA signing verification Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The example loads an Excel workbook, verifies the presence of a signed VBA project, uses reflection to obtain the SignatureInfo and its X509Certificate2, exports the certificate in DER format, converts it to PEM text, and writes the PEM file to disk, with comprehensive error handling.
class ExportVbaCertificate
{
    static void Main()
    {
        // Path to the Excel file that contains the VBA project
        string excelFilePath = "input.xlsx";

        // Path where the PEM file will be saved
        string pemFilePath = "vba_certificate.pem";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine($"Error: The file '{excelFilePath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelFilePath);

            // Check if the workbook has a VBA project and if it is signed
            if (workbook.VbaProject != null && workbook.VbaProject.IsSigned)
            {
                // Try to obtain the SignatureInfo property via reflection (covers different library versions)
                PropertyInfo signatureInfoProp = workbook.VbaProject.GetType().GetProperty("SignatureInfo", BindingFlags.Public | BindingFlags.Instance);
                if (signatureInfoProp != null)
                {
                    object signatureInfo = signatureInfoProp.GetValue(workbook.VbaProject);
                    if (signatureInfo != null)
                    {
                        // Retrieve the Certificate property from the SignatureInfo object
                        PropertyInfo certProp = signatureInfo.GetType().GetProperty("Certificate", BindingFlags.Public | BindingFlags.Instance);
                        X509Certificate2 cert = certProp?.GetValue(signatureInfo) as X509Certificate2;

                        if (cert != null)
                        {
                            // Export the certificate bytes in DER format
                            byte[] derBytes = cert.Export(X509ContentType.Cert);

                            // Convert DER bytes to PEM format
                            string pem = "-----BEGIN CERTIFICATE-----\n" +
                                         Convert.ToBase64String(derBytes, Base64FormattingOptions.InsertLineBreaks) +
                                         "\n-----END CERTIFICATE-----";

                            // Write the PEM string to the specified file
                            File.WriteAllText(pemFilePath, pem);
                            Console.WriteLine($"VBA digital certificate exported successfully to '{pemFilePath}'.");
                        }
                        else
                        {
                            Console.WriteLine("The VBA project is signed, but the certificate could not be retrieved.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Signature information is unavailable for the signed VBA project.");
                    }
                }
                else
                {
                    Console.WriteLine("The current Aspose.Cells version does not expose SignatureInfo; cannot extract the certificate.");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a signed VBA project.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
