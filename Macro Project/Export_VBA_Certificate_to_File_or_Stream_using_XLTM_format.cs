using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    public class ExportVbaCertificate
    {
        public static void Run()
        {
            // Path to an existing macro‑enabled workbook that is already signed.
            // Replace with the actual file path in your environment.
            const string signedWorkbookPath = "SignedWorkbook.xlsm";

            // Verify that the source workbook exists.
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"Source file not found: {signedWorkbookPath}");
                return;
            }

            // Load the signed workbook.
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project associated with the workbook.
            VbaProject vbaProject = workbook.VbaProject;

            // Check whether the VBA project is signed.
            if (!vbaProject.IsSigned)
            {
                Console.WriteLine("The VBA project is not signed. No certificate data available.");
                return;
            }

            // Retrieve the raw certificate data.
            byte[] certData = vbaProject.CertRawData;

            if (certData == null || certData.Length == 0)
            {
                Console.WriteLine("Certificate raw data is empty.");
                return;
            }

            // -----------------------------------------------------------------
            // Export certificate to a physical file.
            // -----------------------------------------------------------------
            const string certFilePath = "VbaCertificate.cer";
            File.WriteAllBytes(certFilePath, certData);
            Console.WriteLine($"Certificate saved to file: {certFilePath}");

            // -----------------------------------------------------------------
            // Export certificate to a memory stream (example of stream usage).
            // -----------------------------------------------------------------
            using (MemoryStream certStream = new MemoryStream())
            {
                certStream.Write(certData, 0, certData.Length);
                // Reset position for any subsequent read operations.
                certStream.Position = 0;

                // Example: write the stream content to another file to verify.
                const string streamCopyPath = "VbaCertificateFromStream.cer";
                using (FileStream fileOut = new FileStream(streamCopyPath, FileMode.Create, FileAccess.Write))
                {
                    certStream.CopyTo(fileOut);
                }

                Console.WriteLine($"Certificate also written from stream to file: {streamCopyPath}");
            }

            // -----------------------------------------------------------------
            // Save the workbook as a macro‑enabled template (XLTM format).
            // -----------------------------------------------------------------
            const string templatePath = "SignedWorkbookTemplate.xltm";
            workbook.Save(templatePath, SaveFormat.Xltm);
            Console.WriteLine($"Workbook saved as macro‑enabled template: {templatePath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificate.Run();
        }
    }
}