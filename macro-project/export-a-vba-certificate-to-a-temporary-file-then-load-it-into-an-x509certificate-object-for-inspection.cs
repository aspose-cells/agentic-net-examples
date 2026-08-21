// Title: Export VBA Project Certificate to a Temporary .cer File and Inspect with X509Certificate2 Using Aspose.Cells for .NET
// Description: Loads a workbook with a signed VBA project, verifies the signature, writes the raw certificate bytes to a uniquely named temporary .cer file, creates an X509Certificate2 object from that file, and displays the certificate's subject, issuer, and validity period.
// Keywords: Aspose.Cells | C# | .NET | VBA project certificate | export certificate | temporary .cer file | X509Certificate2 | certificate inspection | signed macro workbook
// Common Searches: extract VBA signing certificate Aspose.Cells | load VBA project certificate into X509Certificate2 C# | export VBA certificate to .cer file .NET | check VBA macro signature with Aspose.Cells | how to read CertRawData from a signed workbook
// Developer Intent: Retrieve the raw certificate from a signed VBA project, save it as a temporary .cer file, and examine its details with X509Certificate2.
// Use Cases: Verify that a macro‑enabled workbook is signed by a trusted authority before processing. | Log certificate information for compliance auditing of VBA macros. | Compare the extracted certificate thumbprint against an approved list to enforce security policies.
// AI Prompts: Write C# code that uses Aspose.Cells to get CertRawData from a signed VBA project, saves it to a temporary .cer file, and loads it into an X509Certificate2 object for inspection. | Create a reusable method that extracts a VBA project's signing certificate, writes it to a temp file, returns Subject, Issuer, NotBefore, and NotAfter from X509Certificate2. | Explain how to handle unsigned VBA projects or missing certificate data when using Aspose.Cells in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Security.Cryptography.X509Certificates;

// Loads a workbook with a signed VBA project, verifies the signature, writes the raw certificate bytes to a uniquely named temporary .cer file, creates an X509Certificate2 object from that file, and displays the certificate's subject, issuer, and validity period.
class ExportVbaCertificate
{
    static void Main()
    {
        // Load a workbook that contains a signed VBA project
        string workbookPath = "SignedWithVba.xlsm"; // replace with actual path
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Ensure the VBA project is signed and certificate data exists
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Export the raw certificate data to a temporary .cer file
            string tempCertPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".cer");
            File.WriteAllBytes(tempCertPath, vbaProject.CertRawData);
            Console.WriteLine("Certificate exported to temporary file: " + tempCertPath);

            // Load the certificate into an X509Certificate2 object for inspection
            X509Certificate2 certificate = new X509Certificate2(tempCertPath);

            // Display certificate details
            Console.WriteLine("Subject: " + certificate.Subject);
            Console.WriteLine("Issuer: " + certificate.Issuer);
            Console.WriteLine("Valid From: " + certificate.NotBefore);
            Console.WriteLine("Valid To: " + certificate.NotAfter);
        }
        else
        {
            Console.WriteLine("The workbook does not contain a signed VBA project or certificate data is unavailable.");
        }
    }
}
