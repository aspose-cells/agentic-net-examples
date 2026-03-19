using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificateToOxps
{
    static void Main()
    {
        // Load a workbook that contains a signed VBA project
        // (Uses the provided Workbook constructor – a lifecycle rule)
        string inputPath = "SignedWorkbook.xlsm";
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project (provided property)
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project is signed and certificate data exists
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // ----- Export certificate to a file in OXPS format -----
            // The raw certificate bytes are written directly to a file with .oxps extension.
            // This preserves the integrity of the certificate data.
            string certFilePath = "VbaCertificate.oxps";
            File.WriteAllBytes(certFilePath, vbaProject.CertRawData);
            Console.WriteLine($"Certificate exported to file: {certFilePath}");

            // ----- Export certificate to a memory stream -----
            // The same raw bytes are written to a MemoryStream, which can be used for further processing
            // (e.g., sending over a network, attaching to an email, etc.).
            using (MemoryStream certStream = new MemoryStream())
            {
                certStream.Write(vbaProject.CertRawData, 0, vbaProject.CertRawData.Length);
                certStream.Position = 0; // Reset position for any subsequent read operations

                // Example: save the stream content to another OXPS file
                using (FileStream fileStream = new FileStream("VbaCertificateFromStream.oxps", FileMode.Create, FileAccess.Write))
                {
                    certStream.CopyTo(fileStream);
                }
                Console.WriteLine("Certificate exported from stream to file: VbaCertificateFromStream.oxps");
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a signed VBA project or the certificate data is unavailable.");
        }
    }
}