using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    static void Main()
    {
        // Path to the workbook that contains a signed VBA project
        string workbookPath = "SignedWorkbook.xlsm";

        // Path where the PEM file will be saved
        string pemFilePath = "VbaCertificate.pem";

        // Load the workbook (uses Aspose.Cells lifecycle rule)
        Workbook workbook = new Workbook(workbookPath);

        // Get the VBA project from the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project is signed and certificate data exists
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Convert the raw certificate bytes to a Base64 string
            string base64Cert = Convert.ToBase64String(vbaProject.CertRawData);

            // Build the PEM formatted string (64‑character lines)
            StringBuilder pemBuilder = new StringBuilder();
            pemBuilder.AppendLine("-----BEGIN CERTIFICATE-----");
            const int lineLength = 64;
            for (int i = 0; i < base64Cert.Length; i += lineLength)
            {
                int chunkSize = Math.Min(lineLength, base64Cert.Length - i);
                pemBuilder.AppendLine(base64Cert.Substring(i, chunkSize));
            }
            pemBuilder.AppendLine("-----END CERTIFICATE-----");

            // Write the PEM content to a file
            File.WriteAllText(pemFilePath, pemBuilder.ToString());

            Console.WriteLine($"VBA project certificate exported to: {pemFilePath}");
        }
        else
        {
            Console.WriteLine("The workbook does not contain a signed VBA project or the certificate data is unavailable.");
        }
    }
}