// Title: Export VBA Project Certificate to PEM with Aspose.Cells for .NET
// Description: Shows how to load an .xlsm workbook using Aspose.Cells, confirm the VBA project is signed, retrieve the raw certificate via VbaProject.CertRawData, encode it in Base64, wrap it with PEM headers, and write the result to a .pem file for backup or compliance purposes.
// Keywords: Aspose.Cells | C# export VBA certificate | VbaProject CertRawData | PEM file generation | Excel VBA digital signature | certificate backup .NET | signed macro extraction | convert VBA cert to PEM | Aspose.Cells VBA project | certificate to PEM
// Common Searches: export VBA certificate Aspose.Cells | C# extract signed VBA project certificate | save VBA digital signature as PEM file | backup VBA macro certificate .NET | Aspose.Cells VbaProject CertRawData example
// Developer Intent: Backup a signed VBA project's digital certificate by exporting it to a PEM file using Aspose.Cells.
// Use Cases: Create a portable copy of a macro's signing certificate before migrating workbooks. | Provide the PEM certificate to external auditors for macro signature verification. | Store extracted certificates in a secure vault for long‑term compliance tracking.
// AI Prompts: Generate C# code that uses Aspose.Cells to read an .xlsm file, verify the VBA project is signed, and export its certificate to a PEM file. | Write a reusable method that accepts workbook and output paths, handles missing or unsigned VBA projects, and returns a success status. | Provide a PowerShell script that calls a compiled .NET assembly to extract a VBA project's certificate and save it as PEM.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Shows how to load an .xlsm workbook using Aspose.Cells, confirm the VBA project is signed, retrieve the raw certificate via VbaProject.CertRawData, encode it in Base64, wrap it with PEM headers, and write the result to a .pem file for backup or compliance purposes.
class ExportVbaCertificate
{
    static void Main()
    {
        // Path to the workbook that contains a signed VBA project
        string workbookPath = "SignedWorkbook.xlsm";

        // Path where the PEM file will be saved
        string pemPath = "VbaCertificate.pem";

        // Load the workbook (uses Aspose.Cells Workbook load rule)
        Workbook workbook = new Workbook(workbookPath);

        // Get the VBA project from the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Check if the VBA project is signed and certificate data exists
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Retrieve the raw certificate bytes
            byte[] certData = vbaProject.CertRawData;

            // Convert the certificate to Base64 string
            string base64 = Convert.ToBase64String(certData);

            // Build PEM formatted string (64 characters per line)
            StringBuilder pemBuilder = new StringBuilder();
            pemBuilder.AppendLine("-----BEGIN CERTIFICATE-----");
            for (int i = 0; i < base64.Length; i += 64)
            {
                int lineLength = Math.Min(64, base64.Length - i);
                pemBuilder.AppendLine(base64.Substring(i, lineLength));
            }
            pemBuilder.AppendLine("-----END CERTIFICATE-----");

            // Write the PEM content to a file
            File.WriteAllText(pemPath, pemBuilder.ToString());

            Console.WriteLine($"Certificate exported successfully to: {pemPath}");
        }
        else
        {
            Console.WriteLine("The workbook does not contain a signed VBA project or the certificate data is unavailable.");
        }
    }
}
