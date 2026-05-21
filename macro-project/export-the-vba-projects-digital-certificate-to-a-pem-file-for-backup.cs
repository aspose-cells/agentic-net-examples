using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    static void Main()
    {
        // Path to the workbook that contains a signed VBA project
        string workbookPath = "SignedWorkbook.xlsm";

        // Destination path for the exported PEM file
        string pemFilePath = "VbaCertificate.pem";

        // Load the workbook (uses Aspose.Cells lifecycle rule)
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project is signed and certificate data is present
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Create an X509Certificate2 instance from the raw certificate bytes
            X509Certificate2 certificate = new X509Certificate2(vbaProject.CertRawData);

            // Export the certificate to DER format and then encode it as Base64 with line breaks
            string base64Cert = Convert.ToBase64String(
                certificate.Export(X509ContentType.Cert),
                Base64FormattingOptions.InsertLineBreaks);

            // Build the PEM representation
            string pemContent = "-----BEGIN CERTIFICATE-----\n" +
                                base64Cert +
                                "\n-----END CERTIFICATE-----";

            // Write the PEM content to the specified file
            File.WriteAllText(pemFilePath, pemContent);

            Console.WriteLine($"Certificate exported successfully to: {pemFilePath}");
        }
        else
        {
            Console.WriteLine("The VBA project is not signed or certificate data is unavailable.");
        }
    }
}